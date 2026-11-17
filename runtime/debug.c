/* debug.c - the debugger of LinBASIC Studio, inside the program (lbvc --debug).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION).
 *
 * lbvc --debug puts VB_DBG(line) before every statement and a frame (vb_dbgframe) with the
 * variables in every procedure.  When the environment variable VBRT_DEBUG names a Unix socket
 * (a named pipe \\.\pipe\... on Windows: what QLocalServer listens on there), the program connects to the IDE and talks a line protocol (UTF-8, fields split by TAB):
 *
 *   program -> IDE   hello
 *                    break <module> <line> <reason>        reason: bp step pause stop error <n> <text>
 *                    var <name> <type> <value> ... end     (locals)
 *                    frame <module> <procedure> <line> ... end   (call stack, top first)
 *                    ok <value> | err <text>              (eval, exec)
 *   IDE -> program   bp+ <module> <line> | bp- <module> <line> | bpclear
 *                    go | step | over | out | runto <module> <line> | end | pause
 *                    locals [<frame>] | stack | eval <expr> | exec <statement>
 *                    next <module> <line>           Set Next Statement (the procedure that stopped): ok | err
 *                    errgo | errstep                after a break at an error: on with the error (Alt+F5, Alt+F8);
 *                                                   go / step there run the statement again (VB6's F5, F8)
 *                    trap 0|1|2                     Error Trapping: 0 Break on All Errors, 1 Break in Class
 *                                                   Module (the default), 2 Break on Unhandled Errors
 *                    watch <id> <kind> <module> <procedure> <expr>   (TABs) a watch that breaks: kind 1 Break
 *                                                   When Value Is True, 2 Break When Value Changes; module and
 *                                                   procedure empty: all  -> break <module> <line> watch <id>
 *                                                   (Break When Value Is True: when it becomes true); the same
 *                                                   <id> again replaces that watch
 *                    wdel <id> | wclear
 *
 * The program stops (its windows too, as in VB6) while it waits for the IDE. Pause (Ctrl+Break)
 * is read between statements. Immediate: ? expression and assignments, with the variables of the
 * procedure, the module variables, the forms and their controls, Me, members of objects (late
 * binding) and the common functions of VBA.
 */
#include "vbrt_int.h"
#include <ctype.h>
#ifndef _WIN32
#include <unistd.h>
#include <poll.h>
#include <sys/socket.h>
#include <sys/un.h>
#include <signal.h>
#ifndef _WIN32
#include <execinfo.h>
#endif
#endif

int vb_dbg_on;
static int sock = -1;           /* >= 0: connected to the IDE */
static vb_dbgframe *top;
static const char *const *modnames;
static int nmods;
static const vb_dbgvar *gvars;
static int ngvars;
enum { RUN, STEP_INTO, STEP_OVER, STEP_OUT, RUN_TO };
static int mode = RUN, step_depth, runto_mod = -1, runto_line;
static volatile int pause_req;
static unsigned poll_count;
static int err_reported;
static int trap_mode = 1;                       /* 0 Break on All Errors, 1 Break in Class Module, 2 Unhandled */
static const char *modkinds;                    /* C class, U user control, F form, M module (vb_dbg_kinds) */
static int brk_how;                             /* how break_loop ended: 0 go/step, 1 next, 2 errgo/errstep */

static int is_class(int mod)
{
    return modkinds && mod >= 0 && mod < (int)strlen(modkinds) && (modkinds[mod] == 'C' || modkinds[mod] == 'U');
}
void vb_dbg_kinds(const char *kinds) { modkinds = kinds; }

/* ---------------------------------------------------------------- watches that break */
typedef struct { int id, kind, mod; char proc[128]; char expr[1024]; int has; char last[512]; } watch_t;
static watch_t *watches;
static int nwatches, capwatches;

/* ---------------------------------------------------------------- breakpoints */
typedef struct { int mod, line; } bp_t;
static bp_t *bps;
static int nbps, capbps;

static int bp_find(int mod, int line)
{
    for (int i = 0; i < nbps; i++) if (bps[i].mod == mod && bps[i].line == line) return i;
    return -1;
}

static int mod_index(const char *name)
{
    for (int i = 0; i < nmods; i++) if (strcasecmp(modnames[i], name) == 0) return i;
    return -1;
}

static void bp_cmd(const char *cmd)
{
    char mname[256]; int l;
    if (sscanf(cmd + 4, "%255s %d", mname, &l) == 2) {
        int mi = mod_index(mname), i = bp_find(mi, l);
        if (cmd[2] == '+' && mi >= 0 && i < 0) {
            if (nbps == capbps) { capbps = capbps ? capbps * 2 : 16; bps = realloc(bps, capbps * sizeof *bps); }
            bps[nbps].mod = mi; bps[nbps].line = l; nbps++;
        } else if (cmd[2] == '-' && i >= 0) bps[i] = bps[--nbps];
    }
}

/* watch <id> TAB <kind> TAB <module> TAB <procedure> TAB <expression> */
static int watch_find(int id)
{
    for (int i = 0; i < nwatches; i++) if (watches[i].id == id) return i;
    return -1;
}

static void watch_cmd(const char *cmd)
{
    const char *f[5];
    char buf[2048];
    snprintf(buf, sizeof buf, "%s", cmd + 6);
    int n = 0;
    f[n++] = buf;
    for (char *q = buf; *q && n < 5; q++) if (*q == '\t') { *q = 0; f[n++] = q + 1; }
    if (n < 5) return;
    int id = atoi(f[0]), kind = atoi(f[1]), mod = *f[2] ? mod_index(f[2]) : -1;
    int i = watch_find(id);
    if (i >= 0) {                               /* the same id again: the watch was edited */
        watch_t *o = &watches[i];
        if (o->kind == kind && o->mod == mod && !strcmp(o->proc, f[3]) && !strcmp(o->expr, f[4]) && (!*f[2] || mod >= 0))
            return;                             /* nothing changed: it keeps its last value */
        watches[i] = watches[--nwatches];
    }
    if (*f[2] && mod < 0) return;
    if (nwatches == capwatches) {
        capwatches = capwatches ? capwatches * 2 : 8;
        watches = realloc(watches, capwatches * sizeof *watches);
    }
    watch_t *w = &watches[nwatches++];
    memset(w, 0, sizeof *w);
    w->id = id;
    w->kind = kind;
    w->mod = mod;
    snprintf(w->proc, sizeof w->proc, "%s", f[3]);
    snprintf(w->expr, sizeof w->expr, "%s", f[4]);
}

/* what the IDE may send at any time: 1 if cmd was one of these */
static int common_cmd(const char *cmd)
{
    if (!strncmp(cmd, "bp+ ", 4) || !strncmp(cmd, "bp- ", 4)) { bp_cmd(cmd); return 1; }
    if (!strcmp(cmd, "bpclear")) { nbps = 0; return 1; }
    if (!strncmp(cmd, "trap ", 5)) { trap_mode = atoi(cmd + 5); return 1; }
    if (!strncmp(cmd, "watch ", 6)) { watch_cmd(cmd); return 1; }
    if (!strcmp(cmd, "wclear")) { nwatches = 0; return 1; }
    if (!strncmp(cmd, "wdel ", 5)) {
        int i = watch_find(atoi(cmd + 5));
        if (i >= 0) watches[i] = watches[--nwatches];
        return 1;
    }
    return 0;
}

/* ---------------------------------------------------------------- the socket */
#ifdef _WIN32
/* dbgpipe_win.c: the named pipe (windows.h and vbrt.h both name VT_EMPTY ...) */
int vb_dbgpipe_open(const char *path);
long vb_dbgpipe_write(const char *s, size_t n);
long vb_dbgpipe_read(char *s, size_t n);
int vb_dbgpipe_wait(int timeout_ms);
void vb_dbgpipe_close(void);
static int conn_open(const char *path) { return vb_dbgpipe_open(path); }
static long conn_write(const char *s, size_t n) { return vb_dbgpipe_write(s, n); }
static long conn_read(char *s, size_t n) { return vb_dbgpipe_read(s, n); }
static int conn_wait(int timeout_ms) { return vb_dbgpipe_wait(timeout_ms); }
static void conn_close(void) { vb_dbgpipe_close(); }
#else
static int conn_open(const char *path)
{
    int s = socket(AF_UNIX, SOCK_STREAM, 0);
    if (s < 0) return -1;
    struct sockaddr_un a;
    memset(&a, 0, sizeof a);
    a.sun_family = AF_UNIX;
    strncpy(a.sun_path, path, sizeof a.sun_path - 1);
    if (connect(s, (struct sockaddr *)&a, sizeof a) != 0) { close(s); return -1; }
    sock = s;
    return 0;
}
static long conn_write(const char *s, size_t n) { return (long)write(sock, s, n); }
static long conn_read(char *s, size_t n) { return (long)read(sock, s, n); }
static int conn_wait(int timeout_ms)
{
    struct pollfd p = { sock, POLLIN, 0 };
    return poll(&p, 1, timeout_ms) > 0;
}
static void conn_close(void) { if (sock >= 0) close(sock); }
#endif

static void disconnect(void)
{
    conn_close();
    sock = -1;
    vb_dbg_on = 0;
}

static void send_line(const char *s)
{
    if (sock < 0) return;
    size_t n = strlen(s);
    while (n) {
        long k = conn_write(s, n);
        if (k <= 0) { disconnect(); return; }
        s += k; n -= (size_t)k;
    }
    if (conn_write("\n", 1) != 1) disconnect();
}

static void sendf(const char *fmt, ...)
{
    char buf[8192];
    va_list ap;
    va_start(ap, fmt);
    vsnprintf(buf, sizeof buf, fmt, ap);
    va_end(ap);
    send_line(buf);
}

static char inbuf[65536];
static int inlen;

/* one command line; blocks (timeout_ms < 0) or polls */
static int read_line(char *out, int cap, int timeout_ms)
{
    for (;;) {
        char *nl = memchr(inbuf, '\n', inlen);
        if (nl) {
            int n = (int)(nl - inbuf);
            int k = n < cap - 1 ? n : cap - 1;
            memcpy(out, inbuf, k);
            out[k] = 0;
            if (k && out[k - 1] == '\r') out[k - 1] = 0;
            memmove(inbuf, nl + 1, inlen - n - 1);
            inlen -= n + 1;
            return 1;
        }
        if (sock < 0) return 0;
        if (!conn_wait(timeout_ms)) return 0;
        long k = conn_read(inbuf + inlen, sizeof inbuf - inlen - 1);
        if (k <= 0) { disconnect(); return 0; }
        inlen += (int)k;
    }
}

const char *vb_dbg_where(void);
#ifndef _WIN32
/* the robot (VBRT_MONKEY) sends SIGUSR1 to a program that is still running after its time: the program writes where
 * it is (the procedures that are running, the innermost first) and goes on, so a HANG says which loop it was in */
static void where_write(const char *what)
{
    static char buf[2048];
    int n = snprintf(buf, sizeof buf, "monkey: %s at", what);
    vb_dbgframe *f = top;
    for (int k = 0; f && k < 8 && n < (int)sizeof buf - 200; k++, f = f->parent)
        n += snprintf(buf + n, sizeof buf - (size_t)n, "%s %s.%s, line %d", k ? " <" : "",
                      f->mod >= 0 && f->mod < nmods ? modnames[f->mod] : "?", f->proc ? f->proc : "?", (int)f->line);
    if (!top) n += snprintf(buf + n, sizeof buf - (size_t)n, " (no procedure of the program: the runtime or a form)");
    if (n > (int)sizeof buf - 2) n = (int)sizeof buf - 2;
    buf[n++] = '\n';
    ssize_t r = write(2, buf, (size_t)n);
    (void)r;
}
static void where_signal(int sig) { (void)sig; where_write("HANG"); }
/* the robot: a program that crashes (a bug of the runtime or of lbvc) writes where it was, in the program and in the
 * runtime (the addresses: addr2line -f -e prog), then dies of the same signal (EXIT139 ...) */
static void crash_signal(int sig)
{
    char what[64];
    snprintf(what, sizeof what, "CRASH (signal %d, %s)", sig, sig == SIGSEGV ? "SIGSEGV" : sig == SIGABRT ? "SIGABRT" :
             sig == SIGBUS ? "SIGBUS" : sig == SIGFPE ? "SIGFPE" : sig == SIGILL ? "SIGILL" : "?");
    where_write(what);
    void *pc[24];
    int n = backtrace(pc, 24);
    ssize_t r = write(2, "monkey: CRASH native:", 21);
    for (int i = 0; i < n; i++) {
        char a[24];
        int k = snprintf(a, sizeof a, " %p", pc[i]);
        r = write(2, a, (size_t)k);
    }
    r = write(2, "\n", 1);
    (void)r;
    signal(sig, SIG_DFL);
    raise(sig);
}
static void crash_handlers(void)
{
    static char alt[65536];                     /* (a stack of its own: an overflow of the program's stack too) */
    stack_t ss = { 0 };
    ss.ss_sp = alt; ss.ss_size = sizeof alt;
    sigaltstack(&ss, 0);
    struct sigaction sa = { 0 };
    sa.sa_handler = crash_signal;
    sa.sa_flags = SA_ONSTACK | SA_RESETHAND;
    sigemptyset(&sa.sa_mask);
    int sigs[] = { SIGSEGV, SIGABRT, SIGBUS, SIGFPE, SIGILL };
    for (int i = 0; i < 5; i++) sigaction(sigs[i], &sa, 0);
}
#endif
void vb_dbg_init(const char *const *mods, int n, const vb_dbgvar *globals, int ng)
{
    modnames = mods; nmods = n; gvars = globals; ngvars = ng;
    vb_where_hook = vb_dbg_where;
    const char *path = getenv("VBRT_DEBUG");
    if (getenv("VBRT_MONKEY") || getenv("VBRT_TRACE")) vb_dbg_on = 2;   /* only where the program is (errors) */
#ifndef _WIN32
    if (getenv("VBRT_MONKEY")) { signal(SIGUSR1, where_signal); crash_handlers(); }
#endif
    if (!path || !*path) return;
    if (conn_open(path) != 0) return;
#ifdef _WIN32
    sock = 0;
#endif
    vb_dbg_on = 1;
    send_line("hello");
    /* the breakpoints of the IDE come first, then "go" */
    char cmd[4096];
    while (read_line(cmd, sizeof cmd, -1)) {
        if (strcmp(cmd, "go") == 0) break;
        if (strcmp(cmd, "step") == 0) { mode = STEP_INTO; break; }
        if (!strncmp(cmd, "runto ", 6)) {
            char m[256]; int l;
            if (sscanf(cmd + 6, "%255s %d", m, &l) == 2) { runto_mod = mod_index(m); runto_line = l; mode = RUN_TO; }
            break;
        }
        common_cmd(cmd);                        /* bp+, trap, watch ... */
    }
}

/* ---------------------------------------------------------------- values */
static vbvar var_value(const vb_dbgvar *v, int *ok)
{
    *ok = 1;
    if (!v->p) { *ok = 0; return vb_v_empty(); }
    switch (v->vt) {
    case VT_I2: return vb_v_i(*(vb_i16 *)v->p);
    case VT_I4: return vb_v_l(*(vb_i32 *)v->p);
    case VT_R4: return vb_v_f(*(float *)v->p);
    case VT_R8: return vb_v_d(*(double *)v->p);
    case VT_CY: return vb_v_c(*(vb_cy *)v->p);
    case VT_DATE: return vb_v_date(*(vb_date *)v->p);
    case VT_BOOL: return vb_v_bool(*(vb_bool *)v->p);
    case VT_UI1: return vb_v_byte(*(vb_u8 *)v->p);
    case VT_BSTR: return vb_v_s(*(vbs *)v->p);
    case VT_VARIANT: return *(vbvar *)v->p;
    case VT_DISPATCH: return vb_v_obj(*(vbobj *)v->p);
    default:
        if (v->vt & VT_ARRAY) {
            vbarr *a = (vbarr *)v->p;
            if ((v->vt & 0xFFF) == 36) { *ok = 0; return vb_v_empty(); }
            return vb_v_arr(a, v->vt & 0xFFF);
        }
        *ok = 0;
        return vb_v_empty();
    }
}

static const char *vt_name(int vt)
{
    switch (vt & 0xFFF) {
    case VT_I2: return "Integer"; case VT_I4: return "Long"; case VT_R4: return "Single"; case VT_R8: return "Double";
    case VT_CY: return "Currency"; case VT_DATE: return "Date"; case VT_BOOL: return "Boolean"; case VT_UI1: return "Byte";
    case VT_BSTR: return "String"; case VT_VARIANT: return "Variant"; case VT_DISPATCH: return "Object";
    case VT_DECIMAL: return "Decimal";
    case 36: return "(Type)";
    default: return "Variant";
    }
}

/* the text of a value as the Locals window shows it: "text", 12, True, Nothing, Empty */
static void show(vbvar v, char *out, int cap)
{
    if (v.vt & VT_ARRAY) {
        vbarr *a = vb_v_arrptr(v);
        if (!a || a->nd == 0) { snprintf(out, cap, "%s()", vt_name(v.vt)); return; }
        int k = snprintf(out, cap, "%s(", vt_name(v.vt));
        for (int d = 0; d < a->nd && k < cap; d++)
            k += snprintf(out + k, cap - k, "%s%d to %d", d ? ", " : "", a->lb[d], a->ub[d]);
        if (k < cap) snprintf(out + k, cap - k, ")");
        return;
    }
    switch (v.vt) {
    case VT_EMPTY: snprintf(out, cap, "Empty"); return;
    case VT_NULL: snprintf(out, cap, "Null"); return;
    case VT_DISPATCH:
        if (!v.u.o) { snprintf(out, cap, "Nothing"); return; }
        snprintf(out, cap, "%s", vb_stoutf8(vb_typename(v)));
        return;
    case VT_BSTR: {
        const char *s = vb_stoutf8(v.u.s);
        snprintf(out, cap, "\"%s\"", s);
        return;
    }
    case VT_ERROR: snprintf(out, cap, "Error %d", v.u.l); return;
    default:
        snprintf(out, cap, "%s", vb_stoutf8(vb_v2s(v)));
    }
}

static const char *type_of(vbvar v, const vb_dbgvar *d)
{
    if (d && d->vt == VT_VARIANT) {
        static char buf[64];
        snprintf(buf, sizeof buf, "Variant/%s", vb_stoutf8(vb_typename(v)));
        return buf;
    }
    if (d && (d->vt & VT_ARRAY)) return vt_name(d->vt);
    if (d && d->vt == VT_DISPATCH && v.u.o) return vb_stoutf8(vb_typename(v));
    return d ? vt_name(d->vt) : vb_stoutf8(vb_typename(v));
}

/* ---------------------------------------------------------------- errors inside the debugger */
static int take_error(char *msg, int cap)
{
    if (!vb_errpend) return 0;
    vb_err_take(0);
    snprintf(msg, cap, "Run-time error '%d': %s", vb_err_number(), vb_stoutf8(vb_err_description()));
    vb_err_clear();
    return 1;
}

/* ---------------------------------------------------------------- the expression evaluator */
typedef struct { const char *s; int err; char msg[512]; vb_dbgframe *f; } P;

static void skip(P *p) { while (*p->s == ' ' || *p->s == '\t') p->s++; }

static int peekw(P *p, const char *w)
{
    skip(p);
    size_t n = strlen(w);
    return strncasecmp(p->s, w, n) == 0 && !(isalnum((unsigned char)p->s[n]) || p->s[n] == '_');
}

static int eatw(P *p, const char *w) { if (peekw(p, w)) { p->s += strlen(w); return 1; } return 0; }
static int eat(P *p, const char *t) { skip(p); size_t n = strlen(t); if (!strncmp(p->s, t, n)) { p->s += n; return 1; } return 0; }
static void fail(P *p, const char *m) { if (!p->err) { p->err = 1; snprintf(p->msg, sizeof p->msg, "%s", m); } }

static int ident(P *p, char *out, int cap)
{
    skip(p);
    const char *s = p->s;
    if (!(isalpha((unsigned char)*s) || *s == '_' || (unsigned char)*s >= 0x80)) return 0;
    int k = 0;
    while ((isalnum((unsigned char)*s) || *s == '_' || (unsigned char)*s >= 0x80) && k < cap - 2) out[k++] = *s++;
    if (*s == '$' || *s == '%' || *s == '&' || *s == '!' || *s == '#' || *s == '@') out[k++] = *s++;
    out[k] = 0;
    p->s = s;
    return 1;
}

static vbvar expr(P *p);

static int args(P *p, vbvar *argv, int max)
{
    int n = 0;
    if (eat(p, ")")) return 0;
    for (;;) {
        if (n < max) argv[n++] = expr(p); else expr(p);
        if (p->err) return n;
        if (eat(p, ",")) continue;
        if (!eat(p, ")")) fail(p, "Expected: )");
        return n;
    }
}

static const vb_dbgvar *find_var(vb_dbgframe *f, const char *name)
{
    if (f) for (int i = 0; i < f->nvars; i++) if (f->vars[i].name && strcasecmp(f->vars[i].name, name) == 0) return &f->vars[i];
    if (f) for (int i = 0; i < f->nmvars; i++) if (f->mvars[i].name && strcasecmp(f->mvars[i].name, name) == 0) return &f->mvars[i];
    for (int i = 0; i < ngvars; i++) if (strcasecmp(gvars[i].name, name) == 0) return &gvars[i];
    return 0;
}

static void lower(char *d, const char *s, int cap)
{
    int k = 0;
    while (*s && k < cap - 1) { char c = *s++; d[k++] = (char)tolower((unsigned char)c); }
    d[k] = 0;
}

static vbvar call_builtin(P *p, const char *name, int argc, vbvar *a, int *found)
{
    *found = 1;
    char n[64];
    lower(n, name, sizeof n);
    size_t L = strlen(n);
    if (L && n[L - 1] == '$') n[--L] = 0;
#define NEED(k) if (argc < (k)) { fail(p, "Argument not optional"); return vb_v_empty(); }
    if (!strcmp(n, "len")) { NEED(1); return vb_v_l(vb_v2s(a[0]).len); }
    if (!strcmp(n, "ucase") || !strcmp(n, "lcase") || !strcmp(n, "trim") || !strcmp(n, "ltrim") || !strcmp(n, "rtrim")) {
        NEED(1);
        vbs s = vb_v2s(a[0]);
        vbs r = vb_scopy(s);
        if (n[0] == 'u' || (n[0] == 'l' && n[1] == 'c')) {
            for (int i = 0; i < r.len; i++) {
                vb_wchar c = r.p[i];
                r.p[i] = n[0] == 'u' ? (vb_wchar)((c >= 'a' && c <= 'z') || (c >= 0x3B1 && c <= 0x3C9 && c != 0x3C2) ? c - 32 : c)
                                     : (vb_wchar)((c >= 'A' && c <= 'Z') || (c >= 0x391 && c <= 0x3A9) ? c + 32 : c);
            }
            return vb_v_s(r);
        }
        int b = 0, e = r.len;
        if (n[0] != 'r') while (b < e && r.p[b] == ' ') b++;
        if (n[0] != 'l') while (e > b && r.p[e - 1] == ' ') e--;
        r.p += b; r.len = e - b;
        return vb_v_s(r);
    }
    if (!strcmp(n, "left") || !strcmp(n, "right")) {
        NEED(2);
        vbs s = vb_scopy(vb_v2s(a[0]));
        int k = vb_v2l(a[1]);
        if (k < 0) { vb_raise(5); return vb_v_empty(); }
        if (k > s.len) k = s.len;
        if (n[0] == 'r') s.p += s.len - k;
        s.len = k;
        return vb_v_s(s);
    }
    if (!strcmp(n, "mid")) {
        NEED(2);
        vbs s = vb_scopy(vb_v2s(a[0]));
        int st = vb_v2l(a[1]), k = argc > 2 ? vb_v2l(a[2]) : s.len;
        if (st < 1 || k < 0) { vb_raise(5); return vb_v_empty(); }
        if (st > s.len) { s.len = 0; return vb_v_s(s); }
        s.p += st - 1;
        s.len -= st - 1;
        if (k < s.len) s.len = k;
        return vb_v_s(s);
    }
    if (!strcmp(n, "cstr") || !strcmp(n, "str")) {
        NEED(1);
        vbs s = vb_v2s(a[0]);
        if (n[0] == 's' && a[0].vt != VT_BSTR) {
            double d = vb_v2d(a[0]);
            if (d >= 0) { char buf[64]; snprintf(buf, sizeof buf, " %s", vb_stoutf8(s)); return vb_v_s(vb_sfromutf8(buf)); }
        }
        return vb_v_s(s);
    }
    if (!strcmp(n, "cint")) { NEED(1); return vb_v_i(vb_v2i(a[0])); }
    if (!strcmp(n, "clng")) { NEED(1); return vb_v_l(vb_v2l(a[0])); }
    if (!strcmp(n, "cdbl") || !strcmp(n, "val")) {
        NEED(1);
        if (n[0] == 'v') return vb_v_d(atof(vb_stoutf8(vb_v2s(a[0]))));
        return vb_v_d(vb_v2d(a[0]));
    }
    if (!strcmp(n, "csng")) { NEED(1); return vb_v_f(vb_v2f(a[0])); }
    if (!strcmp(n, "cbool")) { NEED(1); return vb_v_bool(vb_v2bool(a[0])); }
    if (!strcmp(n, "cdate")) { NEED(1); return vb_v_date(vb_v2date(a[0])); }
    if (!strcmp(n, "abs")) { NEED(1); return vb_v_d(fabs(vb_v2d(a[0]))); }
    if (!strcmp(n, "int") || !strcmp(n, "fix")) { NEED(1); double d = vb_v2d(a[0]); return vb_v_d(n[0] == 'i' ? floor(d) : trunc(d)); }
    if (!strcmp(n, "sqr")) { NEED(1); return vb_v_d(sqrt(vb_v2d(a[0]))); }
    if (!strcmp(n, "asc")) { NEED(1); vbs s = vb_v2s(a[0]); if (!s.len) { vb_raise(5); return vb_v_empty(); } int c = vb_w2ansi(s.p[0]); return vb_v_i((vb_i16)(c < 0 ? '?' : c)); }
    if (!strcmp(n, "ascw")) { NEED(1); vbs s = vb_v2s(a[0]); if (!s.len) { vb_raise(5); return vb_v_empty(); } return vb_v_i((vb_i16)s.p[0]); }
    if (!strcmp(n, "chr") || !strcmp(n, "chrw")) {
        NEED(1);
        int c = vb_v2l(a[0]);
        vbs s = vb_scopy(vb_sfromutf8(" "));
        s.p[0] = n[3] == 'w' ? (vb_wchar)c : vb_ansi2w((unsigned char)c);
        return vb_v_s(s);
    }
    if (!strcmp(n, "hex")) { NEED(1); char b[32]; snprintf(b, sizeof b, "%X", (unsigned)vb_v2l(a[0])); return vb_v_s(vb_sfromutf8(b)); }
    if (!strcmp(n, "typename")) { NEED(1); return vb_v_s(vb_typename(a[0])); }
    if (!strcmp(n, "isempty")) { NEED(1); return vb_v_bool(a[0].vt == VT_EMPTY ? -1 : 0); }
    if (!strcmp(n, "isnull")) { NEED(1); return vb_v_bool(a[0].vt == VT_NULL ? -1 : 0); }
    if (!strcmp(n, "isnumeric")) {
        NEED(1);
        if (a[0].vt != VT_BSTR) return vb_v_bool((a[0].vt >= VT_I2 && a[0].vt <= VT_CY) || a[0].vt == VT_UI1 ? -1 : 0);
        char *e; const char *s = vb_stoutf8(a[0].u.s); strtod(s, &e);
        return vb_v_bool(*s && !*e ? -1 : 0);
    }
    if (!strcmp(n, "isobject")) { NEED(1); return vb_v_bool(a[0].vt == VT_DISPATCH ? -1 : 0); }
    if (!strcmp(n, "ubound") || !strcmp(n, "lbound")) {
        NEED(1);
        vbarr *ar = vb_v_arrptr(a[0]);
        if (vb_errpend) return vb_v_empty();
        int d = argc > 1 ? vb_v2l(a[1]) : 1;
        return vb_v_l(n[0] == 'u' ? vb_ubound(ar, d) : vb_lbound(ar, d));
    }
    if (!strcmp(n, "instr")) {
        NEED(2);
        int st = 1, i0 = 0;
        if (argc > 2) { st = vb_v2l(a[0]); i0 = 1; }
        vbs h = vb_v2s(a[i0]), nd = vb_v2s(a[i0 + 1]);
        for (int i = st - 1; i + nd.len <= h.len; i++)
            if (!memcmp(h.p + i, nd.p, nd.len * sizeof(vb_wchar))) return vb_v_l(i + 1);
        return vb_v_l(0);
    }
    if (!strcmp(n, "space") || !strcmp(n, "string")) {
        NEED(1);
        int k = vb_v2l(a[0]);
        vb_wchar c = ' ';
        if (n[1] == 't') { NEED(2); vbs cs = vb_v2s(a[1]); c = cs.len ? cs.p[0] : (vb_wchar)vb_v2l(a[1]); }
        if (k < 0) { vb_raise(5); return vb_v_empty(); }
        char *buf = malloc(k + 1); memset(buf, ' ', k); buf[k] = 0;
        vbs s = vb_scopy(vb_sfromutf8(buf)); free(buf);
        for (int i = 0; i < s.len; i++) s.p[i] = c;
        return vb_v_s(s);
    }
    *found = 0;
    return vb_v_empty();
}

/* a name, then .member(args) ... */
static vbvar primary(P *p)
{
    skip(p);
    const char *s = p->s;
    vbvar v = vb_v_empty();
    if (*s == '(') {
        p->s++;
        v = expr(p);
        if (!eat(p, ")")) fail(p, "Expected: )");
    } else if (*s == '"') {
        char buf[4096]; int k = 0;
        s++;
        for (;;) {
            if (!*s) { fail(p, "Expected: \""); break; }
            if (*s == '"') { if (s[1] == '"') { buf[k++] = '"'; s += 2; continue; } s++; break; }
            if (k < (int)sizeof buf - 1) buf[k++] = *s;
            s++;
        }
        buf[k] = 0;
        p->s = s;
        v = vb_v_s(vb_scopy(vb_sfromutf8(buf)));
    } else if (isdigit((unsigned char)*s) || (*s == '.' && isdigit((unsigned char)s[1]))) {
        char *e;
        double d = strtod(s, &e);
        int isint = 1;
        for (const char *q = s; q < e; q++) if (*q == '.' || *q == 'e' || *q == 'E') isint = 0;
        p->s = e;
        if (*p->s == '#' || *p->s == '!' || *p->s == '@' || *p->s == '&' || *p->s == '%') p->s++;
        v = isint && d >= -32768 && d <= 32767 ? vb_v_i((vb_i16)d) : isint && d >= -2147483648.0 && d <= 2147483647.0 ? vb_v_l((vb_i32)d) : vb_v_d(d);
    } else if (*s == '&' && (s[1] == 'H' || s[1] == 'h' || s[1] == 'O' || s[1] == 'o')) {
        char *e;
        long long n = strtoll(s + 2, &e, s[1] == 'H' || s[1] == 'h' ? 16 : 8);
        p->s = e;
        if (*p->s == '&') p->s++;
        v = n >= -32768 && n <= 65535 && (n <= 32767 || 1) ? vb_v_l((vb_i32)n) : vb_v_l((vb_i32)n);
    } else {
        char name[256];
        if (!ident(p, name, sizeof name)) { fail(p, "Expected: expression"); return v; }
        if (!strcasecmp(name, "True")) v = vb_v_bool(-1);
        else if (!strcasecmp(name, "False")) v = vb_v_bool(0);
        else if (!strcasecmp(name, "Nothing")) v = vb_v_obj(0);
        else if (!strcasecmp(name, "Empty")) v = vb_v_empty();
        else if (!strcasecmp(name, "Null")) v = vb_v_null();
        else if (!strcasecmp(name, "vbCrLf")) v = vb_v_s(vb_sfromutf8("\r\n"));
        else if (!strcasecmp(name, "vbTab")) v = vb_v_s(vb_sfromutf8("\t"));
        else if (!strcasecmp(name, "Me")) {
            if (!p->f || !p->f->self) { fail(p, "Invalid use of Me keyword"); return v; }
            v = vb_v_obj(p->f->self);
        } else {
            const vb_dbgvar *dv = find_var(p->f, name);
            vbvar argv[16]; int argc = 0, has_args = 0;
            if (eat(p, "(")) { has_args = 1; argc = args(p, argv, 16); if (p->err) return v; }
            if (dv) {
                int ok;
                v = var_value(dv, &ok);
                if (!ok) { fail(p, "Can't show this variable"); return v; }
                if (has_args) v = (v.vt & VT_ARRAY) ? vb_v_index(v, argc, argv) : vb_late(v, "", VBM_GET, argc, argv);
            } else {
                int found;
                v = call_builtin(p, name, argc, argv, &found);
                if (!found && p->f && p->f->self) {          /* a control or a member of the form/class */
                    char low[256];
                    lower(low, name, sizeof low);
                    v = vb_late_o(p->f->self, low, VBM_GET, argc, argv);
                    if (vb_errpend && vb_err_number_pending() == 438) { vb_err_clear_pending(); found = 0; }
                    else found = 1;
                }
                if (!found) { char m[300]; snprintf(m, sizeof m, "Variable not found: %s", name); fail(p, m); return v; }
            }
        }
    }
    /* .member(args) */
    for (;;) {
        skip(p);
        if (*p->s != '.' || isdigit((unsigned char)p->s[1])) break;
        p->s++;
        char name[256], low[256];
        if (!ident(p, name, sizeof name)) { fail(p, "Expected: identifier"); return v; }
        lower(low, name, sizeof low);
        vbvar argv[16]; int argc = 0;
        if (eat(p, "(")) { argc = args(p, argv, 16); if (p->err) return v; }
        v = vb_late(v, low, VBM_GET, argc, argv);
        if (vb_errpend) return v;
    }
    return v;
}

static vbvar unary(P *p);
static vbvar power(P *p)
{
    vbvar a = primary(p);
    while (!p->err && eat(p, "^")) a = vb_v_pow(a, unary(p));
    return a;
}
static vbvar unary(P *p)
{
    if (eat(p, "-")) return vb_v_neg(unary(p));
    if (eat(p, "+")) return unary(p);
    return power(p);
}
static vbvar mul(P *p)
{
    vbvar a = unary(p);
    for (;;) {
        if (p->err) return a;
        if (eat(p, "*")) a = vb_v_mul(a, unary(p));
        else if (eat(p, "/")) a = vb_v_div(a, unary(p));
        else return a;
    }
}
static vbvar idiv(P *p)
{
    vbvar a = mul(p);
    while (!p->err && eat(p, "\\")) a = vb_v_idiv(a, mul(p));
    return a;
}
static vbvar mod(P *p)
{
    vbvar a = idiv(p);
    while (!p->err && eatw(p, "Mod")) a = vb_v_mod(a, idiv(p));
    return a;
}
static vbvar add(P *p)
{
    vbvar a = mod(p);
    for (;;) {
        if (p->err) return a;
        skip(p);
        if (*p->s == '+') { p->s++; a = vb_v_add(a, mod(p)); }
        else if (*p->s == '-') { p->s++; a = vb_v_sub(a, mod(p)); }
        else return a;
    }
}
static vbvar cat(P *p)
{
    vbvar a = add(p);
    for (;;) {
        skip(p);
        if (p->err || *p->s != '&' || p->s[1] == 'H' || p->s[1] == 'h') return a;
        p->s++;
        a = vb_v_cat(a, add(p));
    }
}
static vbvar cmp(P *p)
{
    vbvar a = cat(p);
    for (;;) {
        if (p->err) return a;
        int op = -1;
        if (eat(p, "<>")) op = 1; else if (eat(p, "<=")) op = 4; else if (eat(p, ">=")) op = 5;
        else if (eat(p, "=")) op = 0; else if (eat(p, "<")) op = 2; else if (eat(p, ">")) op = 3;
        if (op < 0) return a;
        a = vb_v_cmp(a, cat(p), op, 0);
    }
}
static vbvar notx(P *p)
{
    if (eatw(p, "Not")) return vb_v_not(notx(p));
    return cmp(p);
}
static vbvar andx(P *p)
{
    vbvar a = notx(p);
    while (!p->err && eatw(p, "And")) a = vb_v_and(a, notx(p));
    return a;
}
static vbvar expr(P *p)
{
    vbvar a = andx(p);
    for (;;) {
        if (p->err) return a;
        if (eatw(p, "Or")) a = vb_v_or(a, andx(p));
        else if (eatw(p, "Xor")) a = vb_v_xor(a, andx(p));
        else return a;
    }
}

static vb_dbgframe *frame_at(int depth)
{
    vb_dbgframe *f = top;
    while (f && depth-- > 0) f = f->parent;
    return f;
}

static void do_eval(const char *text, int depth)
{
    vb_i32 m = vb_mark();
    P p = { text, 0, "", frame_at(depth) };
    vbvar v = expr(&p);
    char msg[600];
    skip(&p);
    if (!p.err && *p.s) fail(&p, "Expected: end of statement");
    if (take_error(msg, sizeof msg)) sendf("err\t%s", msg);
    else if (p.err) sendf("err\t%s", p.msg);
    else {
        static char out[6000];
        if (v.vt == VT_BSTR) snprintf(out, sizeof out, "%s", vb_stoutf8(v.u.s));
        else show(v, out, sizeof out);
        sendf("ok\t%s", out);
    }
    vb_release(m);
}

/* x = value, Set x = obj, obj.Prop = value, or a method call */
static void do_exec(const char *text, int depth)
{
    vb_i32 m = vb_mark();
    vb_dbgframe *f = frame_at(depth);
    P p = { text, 0, "", f };
    char msg[600];
    int is_set = eatw(&p, "Set");
    eatw(&p, "Let");
    const char *start = p.s;
    /* find a top-level '=' */
    const char *eq = 0;
    int depthp = 0, q = 0;
    for (const char *s = p.s; *s; s++) {
        if (*s == '"') q = !q;
        if (q) continue;
        if (*s == '(') depthp++;
        else if (*s == ')') depthp--;
        else if (*s == '=' && depthp == 0 && s > p.s && s[-1] != '<' && s[-1] != '>' && s[1] != '=') { eq = s; break; }
    }
    if (!eq) {                                 /* a call: Form1.Show, Text1.SetFocus */
        P c = { start, 0, "", f };
        expr(&c);
        if (take_error(msg, sizeof msg)) sendf("err\t%s", msg);
        else if (c.err) sendf("err\t%s", c.msg);
        else send_line("ok\t");
        vb_release(m);
        return;
    }
    P r = { eq + 1, 0, "", f };
    vbvar val = expr(&r);
    if (r.err) { sendf("err\t%s", r.msg); vb_release(m); return; }
    if (take_error(msg, sizeof msg)) { sendf("err\t%s", msg); vb_release(m); return; }
    /* the target: a variable, or obj.member */
    char lhs[1024];
    int n = (int)(eq - start);
    if (n >= (int)sizeof lhs) n = sizeof lhs - 1;
    memcpy(lhs, start, n);
    lhs[n] = 0;
    while (n && lhs[n - 1] == ' ') lhs[--n] = 0;
    char *dot = 0;
    q = 0; depthp = 0;
    for (char *s = lhs; *s; s++) {
        if (*s == '"') q = !q;
        if (q) continue;
        if (*s == '(') depthp++; else if (*s == ')') depthp--;
        else if (*s == '.' && depthp == 0) dot = s;
    }
    if (dot) {
        *dot = 0;
        P o = { lhs, 0, "", f };
        vbvar obj = expr(&o);
        char low[256];
        lower(low, dot + 1, sizeof low);
        if (o.err) { sendf("err\t%s", o.msg); vb_release(m); return; }
        vb_late(obj, low, is_set ? VBM_SET : VBM_LET, 1, &val);
    } else {
        const vb_dbgvar *dv = find_var(f, lhs);
        if (!dv || !dv->p) { sendf("err\tVariable not found: %s", lhs); vb_release(m); return; }
        switch (dv->vt) {
        case VT_I2: *(vb_i16 *)dv->p = vb_v2i(val); break;
        case VT_I4: *(vb_i32 *)dv->p = vb_v2l(val); break;
        case VT_R4: *(float *)dv->p = vb_v2f(val); break;
        case VT_R8: *(double *)dv->p = vb_v2d(val); break;
        case VT_CY: *(vb_cy *)dv->p = vb_v2c(val); break;
        case VT_DATE: *(vb_date *)dv->p = vb_v2date(val); break;
        case VT_BOOL: *(vb_bool *)dv->p = vb_v2bool(val); break;
        case VT_UI1: *(vb_u8 *)dv->p = vb_v2byte(val); break;
        case VT_BSTR: vb_sset((vbs *)dv->p, vb_v2s(val)); break;
        case VT_VARIANT: vb_vset((vbvar *)dv->p, val); break;
        case VT_DISPATCH: vb_oset((vbobj *)dv->p, val.vt == VT_DISPATCH ? val.u.o : vb_v2obj(val)); break;
        default: sendf("err\tCan't assign to this variable"); vb_release(m); return;
        }
    }
    if (take_error(msg, sizeof msg)) sendf("err\t%s", msg);
    else send_line("ok\t");
    vb_release(m);
}

static void send_vars(const char *tag, const vb_dbgvar *vars, int n)
{
    static char val[4000];
    for (int i = 0; i < n; i++) {
        const vb_dbgvar *d = &vars[i];
        if (!d->name || !d->p) continue;
        int ok;
        vbvar v = var_value(d, &ok);
        if (!ok) { sendf("%s\t%s\t%s\t%s", tag, d->name, vt_name(d->vt), d->vt == 36 ? "(Type)" : ""); continue; }
        show(v, val, sizeof val);
        sendf("%s\t%s\t%s\t%s", tag, d->name, type_of(v, d), val);
        if (vb_errpend) vb_err_clear_pending();
    }
}

/* locals: var lines, then mvar lines (the variables of the module, under Me) */
static void do_locals(int depth)
{
    vb_dbgframe *f = frame_at(depth);
    vb_i32 m = vb_mark();
    if (f) {
        send_vars("var", f->vars, f->nvars);
        send_vars("mvar", f->mvars, f->nmvars);
    }
    send_line("end");
    vb_release(m);
}

static void do_stack(void)
{
    for (vb_dbgframe *f = top; f; f = f->parent)
        sendf("frame\t%s\t%s\t%d", f->mod >= 0 && f->mod < nmods ? modnames[f->mod] : "?", f->proc, f->line);
    send_line("end");
}

/* the break mode: commands of the IDE until one of them goes on */
static void break_loop(vb_dbgframe *f, const char *reason)
{
    sendf("break\t%s\t%d\t%s", f->mod >= 0 && f->mod < nmods ? modnames[f->mod] : "?", f->line, reason);
    char cmd[16384];
    mode = RUN;
    runto_mod = -1;
    pause_req = 0;
    brk_how = 0;
    while (sock >= 0) {
        if (!read_line(cmd, sizeof cmd, -1)) break;
        if (!strcmp(cmd, "go")) return;
        if (!strcmp(cmd, "step")) { mode = STEP_INTO; return; }
        if (!strcmp(cmd, "over")) { mode = STEP_OVER; step_depth = f->depth; return; }
        if (!strcmp(cmd, "out")) { mode = STEP_OUT; step_depth = f->depth; return; }
        if (!strcmp(cmd, "errgo")) { brk_how = 2; return; }
        if (!strcmp(cmd, "errstep")) { brk_how = 2; mode = STEP_INTO; return; }
        if (!strcmp(cmd, "end")) { disconnect(); vb_end(); }
        if (!strncmp(cmd, "runto ", 6)) {
            char mname[256]; int l;
            if (sscanf(cmd + 6, "%255s %d", mname, &l) == 2) { runto_mod = mod_index(mname); runto_line = l; mode = RUN_TO; }
            return;
        }
        if (!strncmp(cmd, "next ", 5)) {        /* Set Next Statement: a line of the procedure that stopped */
            char mname[256]; int l, k, st = 0;
            if (f == top && sscanf(cmd + 5, "%255s %d", mname, &l) == 2 && mod_index(mname) == f->mod)
                for (k = 0; k < f->nlines; k++) if (f->lines[2 * k] == l) { st = f->lines[2 * k + 1]; break; }
            if (!st) { send_line("err\tCan't set the next statement to this line: it must be a statement of the "
                                 "procedure that is running"); continue; }
            f->jump = st;
            brk_how = 1;
            mode = STEP_INTO;                   /* it stops again at that line, before running it */
            send_line("ok\t");
            return;
        }
        if (common_cmd(cmd)) continue;
        if (!strncmp(cmd, "locals", 6)) { do_locals(atoi(cmd + 6)); continue; }
        if (!strcmp(cmd, "stack")) { do_stack(); continue; }
        if (!strncmp(cmd, "eval ", 5)) {
            /* eval <depth>\t<expression> */
            char *tab = strchr(cmd + 5, '\t');
            do_eval(tab ? tab + 1 : cmd + 5, tab ? atoi(cmd + 5) : 0);
            continue;
        }
        if (!strncmp(cmd, "exec ", 5)) {
            char *tab = strchr(cmd + 5, '\t');
            do_exec(tab ? tab + 1 : cmd + 5, tab ? atoi(cmd + 5) : 0);
            continue;
        }
    }
}

/* commands that may come while the program runs: pause, breakpoints */
static void poll_commands(void)
{
    char cmd[4096];
    while (read_line(cmd, sizeof cmd, 0)) {
        if (!strcmp(cmd, "pause")) pause_req = 1;
        else if (!strcmp(cmd, "end")) { disconnect(); vb_end(); }
        else common_cmd(cmd);
    }
}

/* where the program is (a --debug build): "Module.Procedure, line N", or "" */
static char err_where[300];                     /* where the error that is not handled came from */
static void where_of(vb_dbgframe *f, char *buf, int cap)
{
    snprintf(buf, (size_t)cap, "%s.%s, line %d", f->mod >= 0 && f->mod < nmods ? modnames[f->mod] : "?",
             f->proc ? f->proc : "?", f->line);
}
const char *vb_dbg_where(void)
{
    static char buf[300];
    if (vb_errpend && err_where[0]) return err_where;
    if (!top) return "";
    where_of(top, buf, sizeof buf);
    return buf;
}

/* ---------------------------------------------------------------- the hooks */
void vb_dbg_enter(vb_dbgframe *f)
{
    f->parent = top;
    f->depth = top ? top->depth + 1 : 0;
    top = f;
}

void vb_dbg_leave(vb_dbgframe *f)
{
    if (top == f) top = f->parent;
    if (mode == STEP_OUT && f->depth <= step_depth) { mode = STEP_INTO; }
}

/* the watches of the procedure that runs: 0, or the reason of a break ("watch\t<id>") */
static const char *check_watches(vb_dbgframe *f)
{
    static char why[40];
    for (int i = 0; i < nwatches; i++) {
        watch_t *w = &watches[i];
        if (w->mod >= 0 && w->mod != f->mod) continue;
        if (w->proc[0] && (!f->proc || strcasecmp(w->proc, f->proc) != 0)) continue;
        vb_i32 m = vb_mark();
        void *st = vb_err_stash();
        P p = { w->expr, 0, "", f };
        vbvar v = expr(&p);
        int bad = p.err || vb_errpend;
        if (vb_errpend) vb_err_clear_pending();   /* (out of context: no break, Err stays as it was) */
        vb_err_unstash(st);
        int hit = 0;
        if (!bad) {
            if (w->kind == 1) {                 /* Break When Value Is True: when it becomes true (not zero) */
                void *s2 = vb_err_stash();
                int now = vb_v2bool(v) != 0;
                if (vb_errpend) { vb_err_clear_pending(); now = 0; }
                vb_err_unstash(s2);
                hit = now && !(w->has && w->last[0] == '1');
                w->last[0] = now ? '1' : '0';
                w->last[1] = 0;
                w->has = 1;
            } else {                            /* Break When Value Changes */
                char now[512];
                show(v, now, sizeof now);
                if (w->has && strcmp(now, w->last) != 0) hit = 1;
                snprintf(w->last, sizeof w->last, "%s", now);
                w->has = 1;
            }
        }
        vb_release(m);
        if (hit) { snprintf(why, sizeof why, "watch\t%d", w->id); return why; }
    }
    return 0;
}

void vb_dbg_line(vb_dbgframe *f, vb_i32 line)
{
    top = f;                                   /* after an error the frames of the callee are gone */
    if (f->jump) return;                       /* Set Next Statement: VB_DBG goes there first */
    f->line = line;
    if (!vb_errpend) err_where[0] = 0;
    if (vb_dbg_on == 2) return;
    if (!vb_errpend) err_reported = 0;
    if ((++poll_count & 1023) == 0 || pause_req) poll_commands();
    const char *why = 0;
    if (pause_req) why = "pause";
    else if (mode == STEP_INTO) why = "step";
    else if (mode == STEP_OVER && f->depth <= step_depth) why = "step";
    else if (mode == RUN_TO && f->mod == runto_mod && line == runto_line) why = "step";
    if (!why && nbps && bp_find(f->mod, line) >= 0) why = "bp";
    if (!why && nwatches) why = check_watches(f);
    if (why) break_loop(f, why);
}

void vb_dbg_stop(vb_dbgframe *f)
{
    if (vb_dbg_on != 1) { vb_stop(); return; }
    break_loop(f, "stop");
}

/* an error that no procedure of the call stack will handle: break at its line, as VB6 does */
/* the break at an error: F5 / F8 run the statement again (VB6), Alt+F5 / Alt+F8 go on with the error (errgo,
 * errstep), Set Next Statement goes to another line; the error is dropped unless it goes on */
static void error_break(vb_dbgframe *f)
{
    char reason[700];
    snprintf(reason, sizeof reason, "error\t%d\t%s", vb_err_number_pending(), vb_err_text_pending());
    err_reported = 1;
    vb_i32 save = vb_errpend;
    vb_errpend = 0;                            /* Locals and Immediate work while the error waits */
    break_loop(f, reason);
    if (brk_how == 2 || sock < 0) { vb_errpend = save; return; }
    if (brk_how == 0 && f->cur && *f->cur > 0) f->jump = *f->cur;   /* the statement again */
    err_reported = 0;
}

void vb_dbg_error(vb_dbgframe *f)
{
    if (vb_dbg_on == 2 && vb_errpend && !err_where[0]) where_of(f, err_where, sizeof err_where);
    if (vb_dbg_on != 1 || !vb_errpend || err_reported) return;
    int cls = is_class(f->mod);
    /* Break on Unhandled Errors: an error of a class module stops at the call, in the caller (MSDN: Debugging Class
     * Modules) - it goes up as the caller's error, and stops there */
    if (trap_mode == 2 && cls && f->parent) return;
    if (trap_mode != 0)
        for (vb_dbgframe *g = f->parent; g; g = g->parent) {
            if (trap_mode == 1 && cls && !is_class(g->mod)) break;   /* Break in Class Module: unhandled in the class */
            if (g->ehmode && *g->ehmode != 0 && !(g->inh && *g->inh)) return;   /* a caller has On Error */
        }
    error_break(f);
}

/* an error that the procedure itself handles (On Error GoTo / Resume Next): only Break on All Errors stops */
void vb_dbg_trap(vb_dbgframe *f)
{
    if (vb_dbg_on != 1 || !vb_errpend || trap_mode != 0 || err_reported) return;   /* (once an error) */
    error_break(f);
}
