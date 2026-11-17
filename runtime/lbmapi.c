/* lbmapi.c - LinBASIC's own provider of Simple MAPI (phase 7.9; lbmapi.h). The messages are in a store folder:
 *   inbox/{cur,new,tmp}   a Maildir: one RFC 822 file per message (MIME: text/plain, multipart/mixed with the
 *                         attachments in base64), the id of a message is its file name up to ":"; in new/ or without
 *                         the flag S in ":2,S" it is unread
 *   sent/cur              what MAPISendMail sent (a copy, the flag S)
 *   tmp                   the attachments MAPIReadMail gives (files, rewritten at each read)
 *   addressbook.txt       "Display Name<TAB>address" lines: what MAPIResolveName and the sending of a recipient
 *                         without an address look up
 *   provider.ini          name=, address= (the originator), profile=, password= (what MAPILogon checks),
 *                         now=YYYY/MM/DD HH:MM (the date of the messages saved or sent: the tests),
 *                         dialog=send|cancel (what the dialogs answer: the address book, a send with MAPI_DIALOG;
 *                         send when not given)
 *                         Linux, how a message sent leaves the computer (after its copy in sent/):
 *                         smtp=host[:port] (SMTP, port 25), else sendmail=path (the default /usr/sbin/sendmail when
 *                         it is there; none: not sent); a send with MAPI_DIALOG (no dialog=) opens the mail program
 *                         of the desktop (xdg-email). A store named by LBMAPI_DIR (the tests) sends by none of them
 *                         unless provider.ini says so.
 *   calls.txt             when it exists: a line for every call with its arguments and its result (the tests)
 * Sessions are slots 1..16 (handle 0x10000 + slot): the same numbers in every run of a test.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "lbmapi.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>
#include <ctype.h>
#include <time.h>
#ifdef _WIN32
#include <windows.h>
#include <direct.h>
#define SEP '\\'
#else
#include <dirent.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <sys/wait.h>
#include <sys/socket.h>
#include <sys/time.h>
#include <netdb.h>
#include <signal.h>
#include <errno.h>
#include <unistd.h>
#define SEP '/'
#endif

/* ---------------------------------------------------------------- small strings */
typedef struct sbuf { char *p; size_t n, cap; } sbuf;
static void sb_putn(sbuf *b, const char *s, size_t n)
{
    if (b->n + n + 1 > b->cap) { b->cap = (b->n + n + 1) * 2 + 64; b->p = (char *)realloc(b->p, b->cap); }
    memcpy(b->p + b->n, s, n);
    b->n += n;
    b->p[b->n] = 0;
}
static void sb_put(sbuf *b, const char *s) { sb_putn(b, s ? s : "", strlen(s ? s : "")); }
static void sb_printf(sbuf *b, const char *fmt, ...)
{
    char tmp[1024];
    va_list ap;
    va_start(ap, fmt);
    int n = vsnprintf(tmp, sizeof tmp, fmt, ap);
    va_end(ap);
    if (n < 0) return;
    if ((size_t)n < sizeof tmp) { sb_putn(b, tmp, (size_t)n); return; }
    char *big = (char *)malloc((size_t)n + 1);
    va_start(ap, fmt);
    vsnprintf(big, (size_t)n + 1, fmt, ap);
    va_end(ap);
    sb_putn(b, big, (size_t)n);
    free(big);
}
static char *xstrdup(const char *s) { s = s ? s : ""; size_t n = strlen(s); char *r = (char *)malloc(n + 1); memcpy(r, s, n + 1); return r; }
static int ieq(const char *a, const char *b)
{
    for (; *a && *b; a++, b++) if (tolower((unsigned char)*a) != tolower((unsigned char)*b)) return 0;
    return *a == *b;
}
static int istarts(const char *s, const char *pre)
{
    for (; *pre; s++, pre++) if (tolower((unsigned char)*s) != tolower((unsigned char)*pre)) return 0;
    return 1;
}
/* strtok_r (not in the C runtime of Windows) */
static char *strtok_r_compat(char *s, const char *delim, char **save)
{
    char *p = s ? s : *save;
    if (!p) return NULL;
    p += strspn(p, delim);
    if (!*p) { *save = p; return NULL; }
    char *e = p + strcspn(p, delim);
    if (*e) *e++ = 0;
    *save = e;
    return p;
}
static char *trim(char *s)
{
    while (*s == ' ' || *s == '\t') s++;
    size_t n = strlen(s);
    while (n && (s[n - 1] == ' ' || s[n - 1] == '\t' || s[n - 1] == '\r' || s[n - 1] == '\n')) s[--n] = 0;
    return s;
}

/* ---------------------------------------------------------------- files */
static void native(char *p)                    /* a path of the program ("C:\x" on Windows, "/a/b\c" on Linux) */
{
#ifndef _WIN32
    for (; *p; p++) if (*p == '\\') *p = '/';
#else
    (void)p;
#endif
}
static FILE *nfopen(const char *path, const char *mode)
{
    char *p = xstrdup(path);
    native(p);
    FILE *f = fopen(p, mode);
    free(p);
    return f;
}
static void mkdir1(const char *p)
{
#ifdef _WIN32
    _mkdir(p);
#else
    mkdir(p, 0755);
#endif
}
static char *join(const char *a, const char *b)
{
    sbuf s = { 0 };
    sb_put(&s, a);
    if (s.n && s.p[s.n - 1] != '/' && s.p[s.n - 1] != '\\') { char c = SEP; sb_putn(&s, &c, 1); }
    sb_put(&s, b);
    return s.p;
}
static char *read_file(const char *path, size_t *len)
{
    FILE *f = nfopen(path, "rb");
    if (!f) return NULL;
    sbuf b = { 0 };
    char tmp[8192];
    size_t n;
    while ((n = fread(tmp, 1, sizeof tmp, f)) > 0) sb_putn(&b, tmp, n);
    fclose(f);
    if (!b.p) b.p = xstrdup("");
    if (len) *len = b.n;
    return b.p;
}
static int write_file(const char *path, const char *data, size_t n)
{
    FILE *f = nfopen(path, "wb");
    if (!f) return -1;
    size_t w = fwrite(data, 1, n, f);
    fclose(f);
    return w == n ? 0 : -1;
}
static int exists(const char *path)
{
    FILE *f = nfopen(path, "rb");
    if (!f) return 0;
    fclose(f);
    return 1;
}
/* the names of the files of a folder (not . ..), count in *n */
static char **list_dir(const char *dir, int *n)
{
    char **out = NULL;
    int k = 0, cap = 0;
#ifdef _WIN32
    char *pat = join(dir, "*");
    WIN32_FIND_DATAA fd;
    HANDLE h = FindFirstFileA(pat, &fd);
    free(pat);
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
            if (k == cap) { cap = cap ? cap * 2 : 16; out = (char **)realloc(out, sizeof(char *) * (size_t)cap); }
            out[k++] = xstrdup(fd.cFileName);
        } while (FindNextFileA(h, &fd));
        FindClose(h);
    }
#else
    DIR *d = opendir(dir);
    if (d) {
        struct dirent *e;
        while ((e = readdir(d))) {
            if (e->d_name[0] == '.') continue;
            if (k == cap) { cap = cap ? cap * 2 : 16; out = (char **)realloc(out, sizeof(char *) * (size_t)cap); }
            out[k++] = xstrdup(e->d_name);
        }
        closedir(d);
    }
#endif
    *n = k;
    return out;
}
static void free_list(char **l, int n) { for (int i = 0; i < n; i++) free(l[i]); free(l); }

/* ---------------------------------------------------------------- the store */
static char store_dir[1024];
static const char *env(const char *name)
{
#ifdef _WIN32
    static char buf[1024];
    DWORD n = GetEnvironmentVariableA(name, buf, sizeof buf);   /* (SetEnvironmentVariable of the program too) */
    return n && n < sizeof buf ? buf : NULL;
#else
    return getenv(name);
#endif
}
const char *lbm_store(void)
{
    const char *e = env("LBMAPI_DIR");
    if (e && *e) snprintf(store_dir, sizeof store_dir, "%s", e);
    else {
#ifdef _WIN32
        snprintf(store_dir, sizeof store_dir, "C:\\LBMAPI\\store");
#else
        const char *x = getenv("XDG_DATA_HOME"), *h = getenv("HOME");
        if (x && *x) snprintf(store_dir, sizeof store_dir, "%s/linbasic-studio/mail", x);
        else snprintf(store_dir, sizeof store_dir, "%s/.local/share/linbasic-studio/mail", h ? h : ".");
#endif
    }
    size_t n = strlen(store_dir);
    while (n > 1 && (store_dir[n - 1] == '/' || store_dir[n - 1] == '\\')) store_dir[--n] = 0;
    return store_dir;
}
static char *spath(const char *rel)            /* a file of the store */
{
    char *r = join(lbm_store(), rel);
#ifdef _WIN32
    for (char *p = r; *p; p++) if (*p == '/') *p = '\\';
#else
    native(r);
#endif
    return r;
}
static void ensure_store(void)
{
    const char *s = lbm_store();
    char *p = xstrdup(s);
    native(p);
    for (char *q = p + 1; *q; q++)              /* the folders above it */
        if (*q == '/' || *q == '\\') { char c = *q; *q = 0; mkdir1(p); *q = c; }
    mkdir1(p);
    free(p);
    static const char *const sub[] = { "inbox", "inbox/cur", "inbox/new", "inbox/tmp", "sent", "sent/cur", "sent/new",
                                       "sent/tmp", "tmp" };
    for (size_t i = 0; i < sizeof sub / sizeof *sub; i++) { char *d = spath(sub[i]); mkdir1(d); free(d); }
}
/* a setting of provider.ini, or NULL */
static char *setting(const char *key)
{
    char *path = spath("provider.ini");
    char *text = read_file(path, NULL);
    free(path);
    if (!text) return NULL;
    char *r = NULL, *save = NULL;
    for (char *line = strtok_r_compat(text, "\n", &save); line; line = strtok_r_compat(NULL, "\n", &save)) {
        char *eq = strchr(line, '=');
        if (!eq || line[0] == ';' || line[0] == '#') continue;
        *eq = 0;
        if (ieq(trim(line), key)) { r = xstrdup(trim(eq + 1)); break; }
    }
    free(text);
    return r;
}

/* ---------------------------------------------------------------- the log of the calls (tests) */
static sbuf logb;
static int logging(void)
{
    char *p = spath("calls.txt");
    int on = exists(p);
    free(p);
    return on;
}
static void lq(const char *s)                  /* a string as the log shows it: "..." or (null), \r \n \t escaped */
{
    if (!s) { sb_put(&logb, "(null)"); return; }
    sb_put(&logb, "\"");
    for (; *s; s++) {
        if (*s == '\r') sb_put(&logb, "\\r");
        else if (*s == '\n') sb_put(&logb, "\\n");
        else if (*s == '\t') sb_put(&logb, "\\t");
        else sb_putn(&logb, s, 1);
    }
    sb_put(&logb, "\"");
}
static void log_recip(const char *what, const lbm_recip *r)
{
    sb_printf(&logb, "  %s", what);
    if (!r) { sb_put(&logb, " (null)\n"); return; }
    sb_printf(&logb, " class=%u name=", (unsigned)r->ulRecipClass);
    lq(r->lpszName);
    sb_put(&logb, " address=");
    lq(r->lpszAddress);
    sb_printf(&logb, " eid=%u\n", (unsigned)r->ulEIDSize);
}
static void log_message(const lbm_message *m)
{
    if (!m) { sb_put(&logb, "  message (null)\n"); return; }
    sb_put(&logb, "  subject="); lq(m->lpszSubject);
    sb_put(&logb, " note="); lq(m->lpszNoteText);
    sb_put(&logb, " type="); lq(m->lpszMessageType);
    sb_put(&logb, "\n  date="); lq(m->lpszDateReceived);
    sb_put(&logb, " conversation="); lq(m->lpszConversationID);
    sb_printf(&logb, " flags=0x%x recips=%u files=%u\n", (unsigned)m->flFlags, (unsigned)m->nRecipCount, (unsigned)m->nFileCount);
    log_recip("originator", m->lpOriginator);
    for (lbm_ulong i = 0; i < m->nRecipCount && m->lpRecips; i++) {
        char w[32];
        snprintf(w, sizeof w, "recip(%u)", (unsigned)i);
        log_recip(w, &m->lpRecips[i]);
    }
    for (lbm_ulong i = 0; i < m->nFileCount && m->lpFiles; i++) {
        const lbm_file *f = &m->lpFiles[i];
        sb_printf(&logb, "  file(%u) flags=%u position=%u path=", (unsigned)i, (unsigned)f->flFlags, (unsigned)f->nPosition);
        lq(f->lpszPathName);
        sb_put(&logb, " name=");
        lq(f->lpszFileName);
        sb_printf(&logb, " type=%s\n", f->lpFileType ? "(set)" : "(null)");
    }
}
static void log_flush(void)
{
    if (!logb.n) return;
    char *p = spath("calls.txt");
    FILE *f = nfopen(p, "ab");
    free(p);
    if (f) {                                   /* CR LF: what Line Input of VB6 reads as lines */
        for (size_t i = 0; i < logb.n; i++) { if (logb.p[i] == '\n') fputc('\r', f); fputc(logb.p[i], f); }
        fclose(f);
    }
    logb.n = 0;
    if (logb.p) logb.p[0] = 0;
}

/* ---------------------------------------------------------------- sessions */
#define LBM_SLOTS 16
#define LBM_BASE 0x10000
static int slot_open[LBM_SLOTS + 1];
static lbm_handle shared_session;              /* the one a logon without MAPI_NEW_SESSION shares */
static int session_ok(lbm_handle s)            /* 0: a temporary session of this call (allowed) */
{
    if (s == 0) return 1;
    return s > LBM_BASE && s <= LBM_BASE + LBM_SLOTS && slot_open[s - LBM_BASE];
}
static const char *ui_str(lbm_handle ui) { return ui ? "hwnd" : "0"; }

/* ---------------------------------------------------------------- arenas (MAPIFreeBuffer) */
typedef struct arena { void *top; void **blocks; int n, cap; struct arena *next; } arena;
static arena *arenas;
static void *aalloc(arena *a, size_t n)
{
    void *p = calloc(1, n ? n : 1);
    if (a->n == a->cap) { a->cap = a->cap ? a->cap * 2 : 16; a->blocks = (void **)realloc(a->blocks, sizeof(void *) * (size_t)a->cap); }
    a->blocks[a->n++] = p;
    return p;
}
static char *astr(arena *a, const char *s)
{
    if (!s) return NULL;
    size_t n = strlen(s);
    char *r = (char *)aalloc(a, n + 1);
    memcpy(r, s, n + 1);
    return r;
}
static arena *anew(void) { arena *a = (arena *)calloc(1, sizeof *a); a->next = arenas; arenas = a; return a; }
lbm_ulong lbm_free_buffer(void *p)
{
    if (!p) return LBM_SUCCESS;
    for (arena **pp = &arenas; *pp; pp = &(*pp)->next) {
        arena *a = *pp;
        if (a->top != p) continue;
        *pp = a->next;
        for (int i = 0; i < a->n; i++) free(a->blocks[i]);
        free(a->blocks);
        free(a);
        return LBM_SUCCESS;
    }
    return LBM_E_FAILURE;
}

/* ---------------------------------------------------------------- base64, quoted-printable, headers */
static const char B64[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
static void b64_encode(sbuf *o, const unsigned char *d, size_t n)
{
    int col = 0;
    for (size_t i = 0; i < n; i += 3) {
        unsigned v = (unsigned)d[i] << 16 | (i + 1 < n ? (unsigned)d[i + 1] << 8 : 0) | (i + 2 < n ? d[i + 2] : 0);
        char q[4] = { B64[v >> 18 & 63], B64[v >> 12 & 63], i + 1 < n ? B64[v >> 6 & 63] : '=', i + 2 < n ? B64[v & 63] : '=' };
        sb_putn(o, q, 4);
        if ((col += 4) >= 76) { sb_put(o, "\r\n"); col = 0; }
    }
    if (col) sb_put(o, "\r\n");
}
static void b64_decode(sbuf *o, const char *s, size_t n)
{
    unsigned v = 0;
    int bits = 0;
    for (size_t i = 0; i < n; i++) {
        const char *p = s[i] ? strchr(B64, s[i]) : NULL;
        if (!p) continue;
        v = v << 6 | (unsigned)(p - B64);
        if ((bits += 6) >= 8) { bits -= 8; char c = (char)(v >> bits & 255); sb_putn(o, &c, 1); }
    }
}
static int hexv(int c) { return c >= '0' && c <= '9' ? c - '0' : c >= 'A' && c <= 'F' ? c - 'A' + 10 : c >= 'a' && c <= 'f' ? c - 'a' + 10 : -1; }
static void qp_decode(sbuf *o, const char *s, size_t n, int header)
{
    for (size_t i = 0; i < n; i++) {
        if (s[i] == '=' && i + 1 < n && (s[i + 1] == '\r' || s[i + 1] == '\n')) {   /* a soft line break */
            i++;
            if (s[i] == '\r' && i + 1 < n && s[i + 1] == '\n') i++;
            continue;
        }
        if (s[i] == '=' && i + 2 < n && hexv(s[i + 1]) >= 0 && hexv(s[i + 2]) >= 0) {
            char c = (char)(hexv(s[i + 1]) * 16 + hexv(s[i + 2]));
            sb_putn(o, &c, 1);
            i += 2;
            continue;
        }
        if (header && s[i] == '_') { sb_put(o, " "); continue; }
        sb_putn(o, s + i, 1);
    }
}
/* =?charset?B|Q?...?= words of a header: their bytes (the charset is not converted) */
static char *decode_words(const char *s)
{
    sbuf o = { 0 };
    sb_put(&o, "");
    while (*s) {
        const char *b = strstr(s, "=?");
        if (!b) { sb_put(&o, s); break; }
        const char *q1 = strchr(b + 2, '?');
        const char *q2 = q1 ? strchr(q1 + 1, '?') : NULL;
        const char *e = q2 ? strstr(q2 + 1, "?=") : NULL;
        if (!e || q2 != q1 + 2) { sb_putn(&o, s, (size_t)(b + 2 - s)); s = b + 2; continue; }
        sb_putn(&o, s, (size_t)(b - s));
        if (toupper((unsigned char)q1[1]) == 'B') b64_decode(&o, q2 + 1, (size_t)(e - q2 - 1));
        else qp_decode(&o, q2 + 1, (size_t)(e - q2 - 1), 1);
        s = e + 2;
        const char *w = s;                     /* the spaces between two encoded words go */
        while (*w == ' ' || *w == '\t' || *w == '\r' || *w == '\n') w++;
        if (w[0] == '=' && w[1] == '?') s = w;
    }
    return o.p;
}

/* ---------------------------------------------------------------- messages: the parts of an RFC 822 file */
typedef struct hdr { char *name, *value; } hdr;
typedef struct part { hdr *h; int nh; char *body; size_t blen; } part;
static void parse_headers(const char *text, size_t len, part *p)
{
    size_t i = 0;
    sbuf cur = { 0 };
    p->h = NULL; p->nh = 0;
    for (;;) {
        size_t e = i;
        while (e < len && text[e] != '\n') e++;
        size_t le = e;
        if (le > i && text[le - 1] == '\r') le--;
        int blank = le == i;
        int cont = !blank && (text[i] == ' ' || text[i] == '\t');
        if (!cont && cur.n) {                  /* the header before is complete */
            char *c = strchr(cur.p, ':');
            if (c) {
                *c = 0;
                p->h = (hdr *)realloc(p->h, sizeof(hdr) * (size_t)(p->nh + 1));
                p->h[p->nh].name = xstrdup(trim(cur.p));
                p->h[p->nh].value = xstrdup(trim(c + 1));
                p->nh++;
            }
            cur.n = 0;
        }
        if (blank || e >= len) { i = e < len ? e + 1 : len; break; }
        if (cont) sb_put(&cur, " ");
        sb_putn(&cur, text + i + (cont ? 1 : 0), le - i - (cont ? 1 : 0));
        i = e + 1;
    }
    free(cur.p);
    p->body = (char *)text + i;
    p->blen = len - i;
}
static const char *hget(const part *p, const char *name)
{
    for (int i = 0; i < p->nh; i++) if (ieq(p->h[i].name, name)) return p->h[i].value;
    return NULL;
}
static void free_part(part *p) { for (int i = 0; i < p->nh; i++) { free(p->h[i].name); free(p->h[i].value); } free(p->h); }
/* a parameter of a header value (boundary=, name=, filename=), or NULL */
static char *hparam(const char *v, const char *key)
{
    if (!v) return NULL;
    size_t kl = strlen(key);
    for (const char *s = strchr(v, ';'); s; s = strchr(s + 1, ';')) {
        const char *k = s + 1;
        while (*k == ' ' || *k == '\t') k++;
        if (!istarts(k, key) || (k[kl] != '=' && k[kl] != ' ')) continue;
        const char *q = strchr(k, '=');
        if (!q) continue;
        q++;
        while (*q == ' ') q++;
        sbuf o = { 0 };
        if (*q == '"') { q++; while (*q && *q != '"') { if (*q == '\\' && q[1]) q++; sb_putn(&o, q++, 1); } }
        else while (*q && *q != ';' && *q != ' ') sb_putn(&o, q++, 1);
        if (!o.p) o.p = xstrdup("");
        char *d = decode_words(o.p);
        free(o.p);
        return d;
    }
    return NULL;
}
static void decode_body(sbuf *o, const part *p)
{
    const char *cte = hget(p, "Content-Transfer-Encoding");
    if (cte && ieq(cte, "base64")) b64_decode(o, p->body, p->blen);
    else if (cte && ieq(cte, "quoted-printable")) qp_decode(o, p->body, p->blen, 0);
    else sb_putn(o, p->body, p->blen);
    if (!o->p) o->p = xstrdup("");
}

/* a message as the provider knows it */
typedef struct addr { lbm_ulong cls; char *name, *address; } addr;
typedef struct att { char *name, *data; size_t len; lbm_ulong pos, flags; } att;
typedef struct msg {
    char *subject, *note, *type, *date, *conv;
    lbm_ulong flags;
    addr orig; int has_orig;
    addr *rc; int nrc;
    att *at; int nat;
} msg;
static void free_msg(msg *m)
{
    free(m->subject); free(m->note); free(m->type); free(m->date); free(m->conv);
    free(m->orig.name); free(m->orig.address);
    for (int i = 0; i < m->nrc; i++) { free(m->rc[i].name); free(m->rc[i].address); }
    free(m->rc);
    for (int i = 0; i < m->nat; i++) { free(m->at[i].name); free(m->at[i].data); }
    free(m->at);
    memset(m, 0, sizeof *m);
}
/* "Name" <a@b>, Name <a@b>, a@b, Name -> name and address ("SMTP:a@b") */
static void parse_addr(const char *s, char **name, char **address)
{
    char *t = xstrdup(s), *v = trim(t);
    char *lt = strrchr(v, '<'), *gt = lt ? strchr(lt, '>') : NULL;
    sbuf a = { 0 };
    sb_put(&a, "");
    if (lt && gt) {
        *gt = 0;
        char *in = trim(lt + 1);
        if (*in) { if (!istarts(in, "SMTP:") && strchr(in, '@')) sb_put(&a, "SMTP:"); sb_put(&a, in); }
        *lt = 0;
        v = trim(v);
        size_t n = strlen(v);
        if (n >= 2 && v[0] == '"' && v[n - 1] == '"') { v[n - 1] = 0; v++; }
        *name = decode_words(v);
    } else if (strchr(v, '@')) {
        if (!istarts(v, "SMTP:")) sb_put(&a, "SMTP:");
        sb_put(&a, v);
        *name = xstrdup(v);
    } else *name = decode_words(v);
    *address = a.p;
    if (*name && !**name && strchr(*address, ':')) { free(*name); *name = xstrdup(strchr(*address, ':') + 1); }
    free(t);
}
static void add_addrs(msg *m, const char *list, lbm_ulong cls)
{
    if (!list) return;
    int q = 0, ang = 0;
    const char *b = list;
    for (const char *s = list;; s++) {
        if (*s == '"') q = !q;
        else if (*s == '<' && !q) ang = 1;
        else if (*s == '>' && !q) ang = 0;
        if (*s == 0 || (*s == ',' && !q && !ang)) {
            char *one = (char *)malloc((size_t)(s - b) + 1);
            memcpy(one, b, (size_t)(s - b));
            one[s - b] = 0;
            if (*trim(one)) {
                m->rc = (addr *)realloc(m->rc, sizeof(addr) * (size_t)(m->nrc + 1));
                m->rc[m->nrc].cls = cls;
                parse_addr(one, &m->rc[m->nrc].name, &m->rc[m->nrc].address);
                m->nrc++;
            }
            free(one);
            if (!*s) break;
            b = s + 1;
        }
    }
}
static const char *const MON[] = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
/* the Date of a message as MAPI gives it: "YYYY/MM/DD HH:MM" */
static char *mapi_date(const char *d)
{
    char out[32] = "";
    if (d) {
        int day = 0, year = 0, hh = 0, mm = 0, mon = -1;
        char mname[8] = "";
        const char *s = strchr(d, ',') ? strchr(d, ',') + 1 : d;
        if (sscanf(s, "%d %7s %d %d:%d", &day, mname, &year, &hh, &mm) >= 3) {
            for (int i = 0; i < 12; i++) if (istarts(mname, MON[i])) mon = i;
            if (mon >= 0) snprintf(out, sizeof out, "%04d/%02d/%02d %02d:%02d", year, mon + 1, day, hh, mm);
        }
    }
    return xstrdup(out);
}
static void add_part(msg *m, const part *p, int top);
static void add_multipart(msg *m, const part *p, const char *ctype)
{
    char *bnd = hparam(ctype, "boundary");
    if (!bnd) return;
    sbuf d = { 0 };
    sb_put(&d, "--");
    sb_put(&d, bnd);
    free(bnd);
    const char *s = p->body, *end = p->body + p->blen;
    const char *cur = NULL;
    while (s < end) {
        const char *ln = s, *e = s;
        while (e < end && *e != '\n') e++;
        size_t ll = (size_t)(e - ln);
        if (ll && ln[ll - 1] == '\r') ll--;
        if (ll >= d.n && !memcmp(ln, d.p, d.n)) {
            if (cur) {                          /* a part ends before this line (its last line break is not its) */
                const char *pe = ln;
                if (pe > cur && pe[-1] == '\n') pe--;
                if (pe > cur && pe[-1] == '\r') pe--;
                part sub;
                parse_headers(cur, (size_t)(pe - cur), &sub);
                add_part(m, &sub, 0);
                free_part(&sub);
            }
            if (ll >= d.n + 2 && ln[d.n] == '-' && ln[d.n + 1] == '-') break;
            cur = e < end ? e + 1 : end;
        }
        s = e < end ? e + 1 : end;
    }
    free(d.p);
}
static void add_part(msg *m, const part *p, int top)
{
    const char *ct = hget(p, "Content-Type");
    const char *disp = hget(p, "Content-Disposition");
    if (ct && istarts(ct, "multipart/")) { add_multipart(m, p, ct); return; }
    char *fname = hparam(disp, "filename");
    if (!fname) fname = hparam(ct, "name");
    int is_att = (disp && istarts(disp, "attachment")) || fname;
    if (!is_att && (!ct || istarts(ct, "text/plain")) && !m->note) {   /* the text of the message */
        sbuf b = { 0 };
        decode_body(&b, p);
        m->note = b.p;
        return;
    }
    if (!is_att && !top) { free(fname); return; }   /* text/html of an alternative ... */
    sbuf b = { 0 };
    decode_body(&b, p);
    m->at = (att *)realloc(m->at, sizeof(att) * (size_t)(m->nat + 1));
    att *a = &m->at[m->nat++];
    a->name = fname ? fname : xstrdup("attachment");
    a->data = b.p;
    a->len = b.n;
    const char *pos = hget(p, "X-MAPI-Position"), *fl = hget(p, "X-MAPI-Flags");
    a->pos = pos ? (lbm_ulong)strtoul(pos, NULL, 10) : (lbm_ulong)0xFFFFFFFFu;
    a->flags = fl ? (lbm_ulong)strtoul(fl, NULL, 10) : 0;
}
static int load_msg(const char *path, msg *m)
{
    size_t len;
    char *text = read_file(path, &len);
    memset(m, 0, sizeof *m);
    if (!text) return -1;
    part p;
    parse_headers(text, len, &p);
    const char *v;
    m->subject = (v = hget(&p, "Subject")) ? decode_words(v) : xstrdup("");
    m->type = xstrdup((v = hget(&p, "X-MAPI-Message-Type")) ? v : "");
    m->conv = xstrdup((v = hget(&p, "X-MAPI-Conversation-ID")) ? v : "");
    m->date = mapi_date(hget(&p, "Date"));
    if (hget(&p, "Disposition-Notification-To") || hget(&p, "Return-Receipt-To")) m->flags |= LBM_RECEIPT_REQUESTED;
    if ((v = hget(&p, "X-MAPI-Sent")) && atoi(v)) m->flags |= LBM_SENT;
    if ((v = hget(&p, "From"))) { m->has_orig = 1; parse_addr(v, &m->orig.name, &m->orig.address); }
    add_addrs(m, hget(&p, "To"), 1);
    add_addrs(m, hget(&p, "Cc"), 2);
    add_addrs(m, hget(&p, "Bcc"), 3);
    add_part(m, &p, 1);
    if (!m->note) m->note = xstrdup("");
    free_part(&p);
    free(text);
    return 0;
}

/* writing: a header value with its address ("SMTP:a@b" -> <a@b>) */
static void put_addr(sbuf *o, const char *name, const char *address)
{
    const char *a = address ? address : "";
    if (istarts(a, "SMTP:")) a += 5;
    if (name && *name && *a && strcmp(name, a)) sb_printf(o, "\"%s\" <%s>", name, a);
    else if (*a) sb_printf(o, "<%s>", a);
    else sb_put(o, name ? name : "");
}
static void put_rfc_date(sbuf *o, time_t t)
{
    struct tm tmv;
#ifdef _WIN32
    localtime_s(&tmv, &t);
#else
    localtime_r(&t, &tmv);
#endif
    static const char *const D[] = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
    sb_printf(o, "%s, %d %s %04d %02d:%02d:%02d", D[tmv.tm_wday], tmv.tm_mday, MON[tmv.tm_mon], tmv.tm_year + 1900,
              tmv.tm_hour, tmv.tm_min, tmv.tm_sec);
}
/* "YYYY/MM/DD HH:MM" (MAPI) -> the Date of RFC 822; 0 when it is not one */
static int date_of_mapi(sbuf *o, const char *d)
{
    int y, mo, da, h = 0, mi = 0;
    if (!d || sscanf(d, "%d/%d/%d %d:%d", &y, &mo, &da, &h, &mi) < 3 || mo < 1 || mo > 12) return 0;
    struct tm t = { 0 };
    t.tm_year = y - 1900; t.tm_mon = mo - 1; t.tm_mday = da; t.tm_hour = h; t.tm_min = mi; t.tm_isdst = -1;
    time_t tt = mktime(&t);
    put_rfc_date(o, tt);
    return 1;
}
static int file_ok(const char *path)           /* an attachment that can be read */
{
    return path && *path && exists(path);
}
/* an RFC 822 file of a MapiMessage; sent: the copy of what was sent; 0 or an error */
static lbm_ulong compose_file(sbuf *o, const lbm_message *m, const msg *orig, int sent)
{
    sbuf ds = { 0 };                           /* its DateReceived, or now (now= of provider.ini: the tests) */
    if (!(m->lpszDateReceived && *m->lpszDateReceived && date_of_mapi(&ds, m->lpszDateReceived))) {
        char *now = setting("now");
        ds.n = 0;
        if (!(now && date_of_mapi(&ds, now))) { ds.n = 0; put_rfc_date(&ds, time(NULL)); }
        free(now);
    }
    sb_put(o, "Date: ");
    sb_putn(o, ds.p, ds.n);
    sb_put(o, "\r\n");
    free(ds.p);
    if (orig && orig->has_orig) { sb_put(o, "From: "); put_addr(o, orig->orig.name, orig->orig.address); sb_put(o, "\r\n"); }
    else if (m->lpOriginator && ((m->lpOriginator->lpszName && *m->lpOriginator->lpszName) ||
                                 (m->lpOriginator->lpszAddress && *m->lpOriginator->lpszAddress))) {
        sb_put(o, "From: "); put_addr(o, m->lpOriginator->lpszName, m->lpOriginator->lpszAddress); sb_put(o, "\r\n");
    }
    static const char *const H[] = { NULL, "To", "Cc", "Bcc" };
    for (lbm_ulong cls = 1; cls <= 3; cls++) {
        int first = 1;
        for (lbm_ulong i = 0; i < m->nRecipCount && m->lpRecips; i++) {
            const lbm_recip *r = &m->lpRecips[i];
            if (r->ulRecipClass != cls) continue;
            sb_put(o, first ? H[cls] : ",\r\n ");
            if (first) sb_put(o, ": ");
            put_addr(o, r->lpszName, r->lpszAddress);
            first = 0;
        }
        if (!first) sb_put(o, "\r\n");
    }
    sb_printf(o, "Subject: %s\r\n", m->lpszSubject ? m->lpszSubject : "");
    if (m->lpszMessageType && *m->lpszMessageType) sb_printf(o, "X-MAPI-Message-Type: %s\r\n", m->lpszMessageType);
    if (m->lpszConversationID && *m->lpszConversationID) sb_printf(o, "X-MAPI-Conversation-ID: %s\r\n", m->lpszConversationID);
    if (m->flFlags & LBM_RECEIPT_REQUESTED) sb_put(o, "Disposition-Notification-To: <>\r\n");
    if (sent || (m->flFlags & LBM_SENT)) sb_put(o, "X-MAPI-Sent: 1\r\n");
    sb_put(o, "MIME-Version: 1.0\r\n");
    const char *note = m->lpszNoteText ? m->lpszNoteText : "";
    if (!m->nFileCount) {
        sb_put(o, "Content-Type: text/plain\r\nContent-Transfer-Encoding: 8bit\r\n\r\n");
        sb_put(o, note);
        return LBM_SUCCESS;
    }
    const char *bnd = "=_lbmapi_part_=";
    sb_printf(o, "Content-Type: multipart/mixed; boundary=\"%s\"\r\n\r\n", bnd);
    sb_printf(o, "--%s\r\nContent-Type: text/plain\r\nContent-Transfer-Encoding: 8bit\r\n\r\n", bnd);
    sb_put(o, note);
    sb_put(o, "\r\n");
    for (lbm_ulong i = 0; i < m->nFileCount; i++) {
        const lbm_file *f = &m->lpFiles[i];
        if (!file_ok(f->lpszPathName)) return LBM_E_ATTACHMENT_NOT_FOUND;
        size_t len;
        char *data = read_file(f->lpszPathName, &len);
        if (!data) return LBM_E_ATTACHMENT_OPEN_FAILURE;
        const char *name = f->lpszFileName && *f->lpszFileName ? f->lpszFileName : f->lpszPathName;
        const char *b1 = strrchr(name, '\\'), *b2 = strrchr(name, '/');
        if (b1 || b2) name = (b1 > b2 ? b1 : b2) + 1;
        sb_printf(o, "--%s\r\nContent-Type: application/octet-stream; name=\"%s\"\r\nContent-Transfer-Encoding: base64\r\n"
                     "Content-Disposition: attachment; filename=\"%s\"\r\n", bnd, name, name);
        if (f->nPosition != 0xFFFFFFFFu) sb_printf(o, "X-MAPI-Position: %u\r\n", (unsigned)f->nPosition);
        if (f->flFlags) sb_printf(o, "X-MAPI-Flags: %u\r\n", (unsigned)f->flFlags);
        sb_put(o, "\r\n");
        b64_encode(o, (const unsigned char *)data, len);
        free(data);
    }
    sb_printf(o, "--%s--\r\n", bnd);
    return LBM_SUCCESS;
}

/* ---------------------------------------------------------------- the messages of the inbox */
typedef struct entry { char *id, *file; int unread; char *date; } entry;
static int cmp_id(const void *a, const void *b) { return strcmp(((const entry *)a)->id, ((const entry *)b)->id); }
static int cmp_date(const void *a, const void *b)
{
    int c = strcmp(((const entry *)a)->date, ((const entry *)b)->date);
    return c ? c : cmp_id(a, b);
}
static char *id_of(const char *fname)
{
    const char *c = strchr(fname, ':');
#ifdef _WIN32
    if (!c) c = strchr(fname, '!');           /* (":" is not allowed in a file name of Windows) */
#endif
    size_t n = c ? (size_t)(c - fname) : strlen(fname);
    char *r = (char *)malloc(n + 1);
    memcpy(r, fname, n);
    r[n] = 0;
    return r;
}
#ifdef _WIN32
#define FLAGSEP "!2,"
#else
#define FLAGSEP ":2,"
#endif
static entry *inbox(int *n)
{
    entry *out = NULL;
    int k = 0;
    static const char *const sub[] = { "inbox/cur", "inbox/new" };
    for (int s = 0; s < 2; s++) {
        char *dir = spath(sub[s]);
        int nf;
        char **files = list_dir(dir, &nf);
        for (int i = 0; i < nf; i++) {
            char *id = id_of(files[i]);
            entry *e = NULL;
            for (int j = 0; j < k && !e; j++)   /* one message for an id: the one in new/ (just delivered) */
                if (!strcmp(out[j].id, id)) { e = &out[j]; free(e->id); free(e->file); free(e->date); }
            if (!e) {
                out = (entry *)realloc(out, sizeof(entry) * (size_t)(k + 1));
                e = &out[k++];
            }
            e->id = id;
            e->file = join(dir, files[i]);
            const char *fl = strstr(files[i], FLAGSEP);
            e->unread = s == 1 || !fl || !strchr(fl + 3, 'S');
            msg m;
            e->date = load_msg(e->file, &m) == 0 ? xstrdup(m.date) : xstrdup("");
            free_msg(&m);
        }
        free_list(files, nf);
        free(dir);
    }
    *n = k;
    return out;
}
static void free_inbox(entry *e, int n) { for (int i = 0; i < n; i++) { free(e[i].id); free(e[i].file); free(e[i].date); } free(e); }
static int find_entry(entry *e, int n, const char *id)
{
    for (int i = 0; i < n; i++) if (!strcmp(e[i].id, id)) return i;
    return -1;
}
static char *next_id(void)                     /* the next free number of the store */
{
    long best = 0;
    static const char *const sub[] = { "inbox/cur", "inbox/new", "sent/cur", "sent/new" };
    for (int s = 0; s < 4; s++) {
        char *dir = spath(sub[s]);
        int nf;
        char **files = list_dir(dir, &nf);
        for (int i = 0; i < nf; i++) { long v = strtol(files[i], NULL, 10); if (v > best) best = v; }
        free_list(files, nf);
        free(dir);
    }
    char buf[32];
    snprintf(buf, sizeof buf, "%06ld", best + 1);
    return xstrdup(buf);
}

/* ---------------------------------------------------------------- the address book */
/* the entries that match a name: exact first (any number of them), else those that begin with it */
static int lookup(const char *name, char ***names, char ***addrs)
{
    char *path = spath("addressbook.txt");
    char *text = read_file(path, NULL);
    free(path);
    *names = *addrs = NULL;
    if (!text) return 0;
    int n = 0, exact = 0;
    for (int pass = 0; pass < 2 && !n; pass++) {
        char *copy = xstrdup(text), *save = NULL;
        for (char *line = strtok_r_compat(copy, "\n", &save); line; line = strtok_r_compat(NULL, "\n", &save)) {
            char *tab = strchr(line, '\t');
            if (!tab) continue;
            *tab = 0;
            char *nm = trim(line), *ad = trim(tab + 1);
            if (pass == 0 ? !ieq(nm, name) : !istarts(nm, name)) continue;
            *names = (char **)realloc(*names, sizeof(char *) * (size_t)(n + 1));
            *addrs = (char **)realloc(*addrs, sizeof(char *) * (size_t)(n + 1));
            (*names)[n] = xstrdup(nm);
            (*addrs)[n] = xstrdup(ad);
            n++;
            exact = pass == 0;
        }
        free(copy);
    }
    (void)exact;
    free(text);
    return n;
}
static void free_lookup(char **a, char **b, int n) { for (int i = 0; i < n; i++) { free(a[i]); free(b[i]); } free(a); free(b); }
static int dialog_sends(void)                  /* what a dialog of the provider answers */
{
    char *d = setting("dialog");
    int r = !d || !ieq(d, "cancel");
    free(d);
    return r;
}
/* a name -> its name and address: 0, LBM_E_UNKNOWN_RECIPIENT or LBM_E_AMBIGUOUS_RECIPIENT */
static lbm_ulong resolve(const char *name, int dialog, char **rname, char **raddr)
{
    char **ns, **as;
    int n = lookup(name ? name : "", &ns, &as);
    lbm_ulong r = LBM_SUCCESS;
    if (n == 1 || (n > 1 && dialog && dialog_sends())) { *rname = xstrdup(ns[0]); *raddr = xstrdup(as[0]); }
    else if (n > 1) r = LBM_E_AMBIGUOUS_RECIPIENT;
    else if (name && strchr(name, '@')) {
        *rname = xstrdup(istarts(name, "SMTP:") ? name + 5 : name);
        sbuf a = { 0 };
        if (!istarts(name, "SMTP:")) sb_put(&a, "SMTP:");
        sb_put(&a, name);
        *raddr = a.p;
    } else r = LBM_E_UNKNOWN_RECIPIENT;
    free_lookup(ns, as, n);
    return r;
}

/* ---------------------------------------------------------------- sending (Linux) */
#ifndef _WIN32
static const char *bare(const char *a) { return !a ? "" : istarts(a, "SMTP:") ? a + 5 : a; }
static int own_store(void) { const char *e = getenv("LBMAPI_DIR"); return e && *e; }   /* the tests */
/* a program with its arguments (no shell), input on its standard input: its exit status, 127 not there */
static int run(char *const argv[], const char *input, size_t n)
{
    int fd[2];
    if (pipe(fd)) return -1;
    void (*old)(int) = signal(SIGPIPE, SIG_IGN);
    pid_t pid = fork();
    if (pid < 0) { close(fd[0]); close(fd[1]); signal(SIGPIPE, old); return -1; }
    if (pid == 0) {
        dup2(fd[0], 0);
        close(fd[0]); close(fd[1]);
        execvp(argv[0], argv);
        _exit(127);
    }
    close(fd[0]);
    for (size_t off = 0; input && off < n;) {
        ssize_t w = write(fd[1], input + off, n - off);
        if (w <= 0) break;
        off += (size_t)w;
    }
    close(fd[1]);
    int st = 0;
    while (waitpid(pid, &st, 0) < 0 && errno == EINTR) {}
    signal(SIGPIPE, old);
    return WIFEXITED(st) ? WEXITSTATUS(st) : -1;
}
/* the message as it goes out: no Bcc header (its recipients are given to the mail system) */
static lbm_ulong wire(sbuf *o, const lbm_message *m)
{
    lbm_message c = *m;
    lbm_recip *rs = (lbm_recip *)calloc(m->nRecipCount ? m->nRecipCount : 1, sizeof(lbm_recip));
    lbm_ulong k = 0;
    for (lbm_ulong i = 0; i < m->nRecipCount; i++) if (m->lpRecips[i].ulRecipClass != 3) rs[k++] = m->lpRecips[i];
    c.lpRecips = rs;
    c.nRecipCount = k;
    msg from = { 0 };
    char *nm = setting("name"), *ad = setting("address");
    if ((nm && *nm) || (ad && *ad)) { from.has_orig = 1; from.orig.name = nm ? nm : xstrdup(""); from.orig.address = ad ? ad : xstrdup(""); nm = ad = NULL; }
    free(nm); free(ad);
    lbm_ulong r = compose_file(o, &c, &from, 1);
    free_msg(&from);
    free(rs);
    return r;
}
static lbm_ulong by_sendmail(const char *path, const lbm_message *m)
{
    sbuf o = { 0 };
    lbm_ulong r = wire(&o, m);
    if (r) { free(o.p); return r; }
    char **argv = (char **)calloc(m->nRecipCount + 4, sizeof(char *));
    int k = 0;
    argv[k++] = (char *)path;
    argv[k++] = (char *)"-i";
    argv[k++] = (char *)"--";
    for (lbm_ulong i = 0; i < m->nRecipCount; i++) argv[k++] = (char *)bare(m->lpRecips[i].lpszAddress);
    int st = run(argv, o.p ? o.p : "", o.n);
    free(argv);
    free(o.p);
    return st == 0 ? LBM_SUCCESS : st == 127 ? LBM_E_NOT_SUPPORTED : LBM_E_NETWORK_FAILURE;
}
/* SMTP (RFC 5321, without TLS): the reply code of the server, -1 when it is gone */
static int smtp_reply(int s)
{
    char line[1024];
    size_t n = 0;
    for (;;) {
        char ch;
        ssize_t g = recv(s, &ch, 1, 0);
        if (g <= 0) return -1;
        if (ch == '\n') {
            line[n] = 0;
            if (n >= 4 && line[3] == '-') { n = 0; continue; }     /* a line of a reply of many */
            return atoi(line);
        }
        if (n < sizeof line - 1) line[n++] = ch;
    }
}
static int smtp_cmd(int s, const char *cmd, const char *arg)
{
    sbuf c = { 0 };
    sb_put(&c, cmd);
    if (arg) sb_put(&c, arg);
    sb_put(&c, "\r\n");
    ssize_t w = send(s, c.p, c.n, 0);
    free(c.p);
    return w < 0 ? -1 : smtp_reply(s);
}
static lbm_ulong by_smtp(const char *server, const lbm_message *m)
{
    char host[256], port[16] = "25";
    snprintf(host, sizeof host, "%s", server);
    char *colon = strrchr(host, ':');
    if (colon) { snprintf(port, sizeof port, "%s", colon + 1); *colon = 0; }
    sbuf o = { 0 };
    lbm_ulong r = wire(&o, m);
    if (r) { free(o.p); return r; }
    struct addrinfo hints = { 0 }, *res = NULL;
    hints.ai_socktype = SOCK_STREAM;
    int s = -1;
    if (!getaddrinfo(host, port, &hints, &res)) {
        for (struct addrinfo *a = res; a && s < 0; a = a->ai_next) {
            s = socket(a->ai_family, a->ai_socktype, a->ai_protocol);
            if (s < 0) continue;
            struct timeval tv = { 30, 0 };
            setsockopt(s, SOL_SOCKET, SO_RCVTIMEO, &tv, sizeof tv);
            setsockopt(s, SOL_SOCKET, SO_SNDTIMEO, &tv, sizeof tv);
            if (connect(s, a->ai_addr, a->ai_addrlen)) { close(s); s = -1; }
        }
        freeaddrinfo(res);
    }
    r = LBM_E_NETWORK_FAILURE;
    if (s >= 0) {
        char *ad = setting("address");
        sbuf from = { 0 };
        sb_printf(&from, "<%s>", bare(ad));
        free(ad);
        int ok = smtp_reply(s) == 220;
        if (ok && smtp_cmd(s, "EHLO ", "localhost") != 250) ok = smtp_cmd(s, "HELO ", "localhost") == 250;
        if (ok) ok = smtp_cmd(s, "MAIL FROM:", from.p) == 250;
        for (lbm_ulong i = 0; ok && i < m->nRecipCount; i++) {
            sbuf to = { 0 };
            sb_printf(&to, "<%s>", bare(m->lpRecips[i].lpszAddress));
            int c = smtp_cmd(s, "RCPT TO:", to.p);
            ok = c == 250 || c == 251;
            free(to.p);
        }
        if (ok) ok = smtp_cmd(s, "DATA", NULL) == 354;
        if (ok) {                              /* the lines, a "." at the start of one doubled */
            sbuf d = { 0 };
            const char *p = o.p ? o.p : "";
            int bol = 1;
            for (; *p; p++) {
                if (bol && *p == '.') sb_put(&d, ".");
                if (*p == '\n' && (p == o.p || p[-1] != '\r')) sb_put(&d, "\r");
                sb_putn(&d, p, 1);
                bol = *p == '\n';
            }
            if (!bol) sb_put(&d, "\r\n");
            sb_put(&d, ".\r\n");
            ok = send(s, d.p, d.n, 0) == (ssize_t)d.n && smtp_reply(s) == 250;
            free(d.p);
        }
        smtp_cmd(s, "QUIT", NULL);
        close(s);
        free(from.p);
        if (ok) r = LBM_SUCCESS;
    }
    free(o.p);
    return r;
}
/* a message sent (its recipients resolved) out of the computer, as provider.ini says */
static lbm_ulong deliver(const lbm_message *m)
{
    char *smtp = setting("smtp"), *sm = setting("sendmail");
    lbm_ulong r = LBM_SUCCESS;
    if (smtp && *smtp) r = by_smtp(smtp, m);
    else if (sm && *sm) { if (!ieq(sm, "none")) r = by_sendmail(sm, m); }
    else if (!own_store() && exists("/usr/sbin/sendmail")) r = by_sendmail("/usr/sbin/sendmail", m);
    free(smtp); free(sm);
    return r;
}
/* a send with MAPI_DIALOG: the mail program of the desktop, with the message in it (xdg-email) */
static int desktop_dialog(void)
{
    char *d = setting("dialog");
    int r = !d && !own_store();
    free(d);
    return r;
}
static lbm_ulong xdg_email(const lbm_message *m)
{
    char **argv = (char **)calloc(2 * (m->nRecipCount + m->nFileCount) + 8, sizeof(char *));
    char **own = (char **)calloc(m->nFileCount + 1, sizeof(char *));
    int k = 0;
    argv[k++] = (char *)"xdg-email";
    if (m->lpszSubject && *m->lpszSubject) { argv[k++] = (char *)"--subject"; argv[k++] = m->lpszSubject; }
    if (m->lpszNoteText && *m->lpszNoteText) { argv[k++] = (char *)"--body"; argv[k++] = m->lpszNoteText; }
    for (lbm_ulong i = 0; i < m->nRecipCount; i++) {
        const lbm_recip *r = &m->lpRecips[i];
        const char *a = r->lpszAddress && *r->lpszAddress ? bare(r->lpszAddress) : r->lpszName ? r->lpszName : "";
        if (r->ulRecipClass == 2) argv[k++] = (char *)"--cc";
        else if (r->ulRecipClass == 3) argv[k++] = (char *)"--bcc";
        else continue;
        argv[k++] = (char *)a;
    }
    for (lbm_ulong i = 0; i < m->nFileCount; i++) {
        own[i] = xstrdup(m->lpFiles[i].lpszPathName);
        native(own[i]);
        argv[k++] = (char *)"--attach";
        argv[k++] = own[i];
    }
    for (lbm_ulong i = 0; i < m->nRecipCount; i++) {
        const lbm_recip *r = &m->lpRecips[i];
        if (r->ulRecipClass == 2 || r->ulRecipClass == 3) continue;
        argv[k++] = (char *)(r->lpszAddress && *r->lpszAddress ? bare(r->lpszAddress) : r->lpszName ? r->lpszName : "");
    }
    int st = run(argv, NULL, 0);
    for (lbm_ulong i = 0; i < m->nFileCount; i++) free(own[i]);
    free(own);
    free(argv);
    return st == 0 ? LBM_SUCCESS : st == 127 ? LBM_E_NOT_SUPPORTED : LBM_E_FAILURE;
}
#endif

/* ---------------------------------------------------------------- the functions */
lbm_ulong lbm_logon(lbm_handle ui, const char *profile, const char *password, lbm_ulong flags, lbm_handle *session)
{
    ensure_store();
    int lg = logging();
    lbm_ulong r = LBM_SUCCESS;
    lbm_handle h = 0;
    char *wp = setting("profile"), *ww = setting("password");
    if (wp && *wp && profile && *profile && !ieq(wp, profile)) r = LBM_E_LOGIN_FAILURE;
    else if (ww && *ww && (!password || strcmp(ww, password)) && !(flags & (LBM_LOGON_UI | LBM_PASSWORD_UI)))
        r = LBM_E_LOGIN_FAILURE;
    else if (ww && *ww && (!password || strcmp(ww, password)) && !dialog_sends()) r = LBM_USER_ABORT;
    else if (!(flags & LBM_NEW_SESSION) && shared_session && session_ok(shared_session) && shared_session) h = shared_session;
    else {
        int s = 1;
        while (s <= LBM_SLOTS && slot_open[s]) s++;
        if (s > LBM_SLOTS) r = LBM_E_TOO_MANY_SESSIONS;
        else {
            slot_open[s] = 1;
            h = LBM_BASE + (lbm_handle)s;
            if (!(flags & LBM_NEW_SESSION) || !shared_session) shared_session = h;
        }
    }
    free(wp); free(ww);
    if (session) *session = r == LBM_SUCCESS ? h : 0;
    if (lg) {
        sb_printf(&logb, "MAPILogon(ui=%s, profile=", ui_str(ui));
        lq(profile);
        sb_put(&logb, ", password=");
        lq(password);
        sb_printf(&logb, ", flags=0x%x) = %u session=0x%lx\n", (unsigned)flags, (unsigned)r, (unsigned long)(r ? 0 : h));
        log_flush();
    }
    return r;
}

lbm_ulong lbm_logoff(lbm_handle session, lbm_handle ui, lbm_ulong flags)
{
    lbm_ulong r = LBM_SUCCESS;
    if (!session || !session_ok(session)) r = LBM_E_INVALID_SESSION;
    else {
        slot_open[session - LBM_BASE] = 0;
        if (shared_session == session) shared_session = 0;
    }
    if (logging()) {
        sb_printf(&logb, "MAPILogoff(session=0x%lx, ui=%s, flags=0x%x) = %u\n", (unsigned long)session, ui_str(ui),
                  (unsigned)flags, (unsigned)r);
        log_flush();
    }
    return r;
}

lbm_ulong lbm_send_mail(lbm_handle session, lbm_handle ui, const lbm_message *m, lbm_ulong flags)
{
    ensure_store();
    lbm_ulong r = LBM_SUCCESS;
    int lg = logging();
    if (lg) {
        sb_printf(&logb, "MAPISendMail(session=0x%lx, ui=%s, flags=0x%x)\n", (unsigned long)session, ui_str(ui), (unsigned)flags);
        log_message(m);
    }
    lbm_message copy;
    lbm_recip *rc = NULL;
    char **keep = NULL;
    int nkeep = 0;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    else if (!m) r = LBM_E_FAILURE;
    else if ((flags & LBM_DIALOG) && !dialog_sends()) r = LBM_USER_ABORT;
    else if (!m->nRecipCount && !(flags & LBM_DIALOG)) r = LBM_E_INVALID_RECIPS;
#ifndef _WIN32
    if (!r && (flags & LBM_DIALOG) && desktop_dialog()) {   /* the mail program of the desktop sends it */
        for (lbm_ulong i = 0; i < m->nFileCount && !r; i++)
            if (!file_ok(m->lpFiles[i].lpszPathName)) r = LBM_E_ATTACHMENT_NOT_FOUND;
        if (!r) r = xdg_email(m);
        if (lg) { sb_printf(&logb, "  = %u\n", (unsigned)r); log_flush(); }
        return r;
    }
#endif
    if (!r) {                                  /* the recipients, resolved */
        copy = *m;
        rc = (lbm_recip *)calloc(m->nRecipCount ? m->nRecipCount : 1, sizeof(lbm_recip));
        for (lbm_ulong i = 0; i < m->nRecipCount && !r; i++) {
            rc[i] = m->lpRecips[i];
            if (rc[i].ulRecipClass < 1 || rc[i].ulRecipClass > 3) { r = LBM_E_BAD_RECIPTYPE; break; }
            if (rc[i].lpszAddress && *rc[i].lpszAddress) continue;
            char *n = NULL, *a = NULL;
            r = resolve(rc[i].lpszName, (flags & LBM_DIALOG) != 0, &n, &a);
            if (r) break;
            keep = (char **)realloc(keep, sizeof(char *) * (size_t)(nkeep + 2));
            keep[nkeep++] = n; keep[nkeep++] = a;
            rc[i].lpszName = n; rc[i].lpszAddress = a;
        }
        copy.lpRecips = rc;
        for (lbm_ulong i = 0; i < m->nFileCount && !r; i++)
            if (!file_ok(m->lpFiles[i].lpszPathName)) r = LBM_E_ATTACHMENT_NOT_FOUND;
    }
    if (!r) {                                  /* the copy in sent/ */
        msg from = { 0 };
        char *nm = setting("name"), *ad = setting("address");
        if ((nm && *nm) || (ad && *ad)) { from.has_orig = 1; from.orig.name = nm ? nm : xstrdup(""); from.orig.address = ad ? ad : xstrdup(""); nm = ad = NULL; }
        free(nm); free(ad);
        sbuf o = { 0 };
        r = compose_file(&o, &copy, &from, 1);
        free_msg(&from);
        if (!r) {
            char *id = next_id();
            sbuf rel = { 0 };
            sb_printf(&rel, "sent/cur/%s" FLAGSEP "S", id);
            char *path = spath(rel.p);
            if (write_file(path, o.p ? o.p : "", o.n)) r = LBM_E_DISK_FULL;
            free(path); free(rel.p); free(id);
        }
        free(o.p);
#ifndef _WIN32
        if (!r) r = deliver(&copy);            /* out of the computer */
#endif
    }
    free(rc);
    for (int i = 0; i < nkeep; i++) free(keep[i]);
    free(keep);
    if (lg) { sb_printf(&logb, "  = %u\n", (unsigned)r); log_flush(); }
    return r;
}

lbm_ulong lbm_send_documents(lbm_handle ui, const char *delims, const char *paths, const char *names)
{
    if (logging()) {
        sb_printf(&logb, "MAPISendDocuments(ui=%s, delims=", ui_str(ui));
        lq(delims); sb_put(&logb, ", paths="); lq(paths); sb_put(&logb, ", names="); lq(names);
        sb_put(&logb, ") = 0\n");
        log_flush();
    }
    return LBM_SUCCESS;
}

lbm_ulong lbm_find_next(lbm_handle session, lbm_handle ui, const char *type, const char *seed, lbm_ulong flags, char *id)
{
    ensure_store();
    lbm_ulong r = LBM_SUCCESS;
    if (id) id[0] = 0;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    int n = 0;
    entry *e = r ? NULL : inbox(&n);
    if (!r) {
        qsort(e, (size_t)n, sizeof *e, (flags & LBM_GUARANTEE_FIFO) ? cmp_date : cmp_id);
        int start = 0;
        if (seed && *seed) {
            int i = find_entry(e, n, seed);
            if (i < 0) r = LBM_E_INVALID_MESSAGE; else start = i + 1;
        }
        int found = -1;
        for (int i = start; !r && i < n; i++) {
            if ((flags & LBM_UNREAD_ONLY) && !e[i].unread) continue;
            if (type && *type) {               /* the messages of that class or of its subclasses (IPM: all of
                                                  them, IPM.Note: those without a type of their own) */
                msg m;
                int ok = load_msg(e[i].file, &m) == 0;
                const char *mt = ok && *m.type ? m.type : "IPM.Note";
                size_t tl = strlen(type);
                ok = ok && (ieq(mt, type) || (istarts(mt, type) && mt[tl] == '.'));
                free_msg(&m);
                if (!ok) continue;
            }
            found = i;
            break;
        }
        if (!r && found < 0) r = LBM_E_NO_MESSAGES;
        if (!r && id) snprintf(id, 512, "%s", e[found].id);
    }
    free_inbox(e, n);
    if (logging()) {
        sb_printf(&logb, "MAPIFindNext(session=0x%lx, ui=%s, type=", (unsigned long)session, ui_str(ui));
        lq(type); sb_put(&logb, ", seed="); lq(seed);
        sb_printf(&logb, ", flags=0x%x) = %u id=", (unsigned)flags, (unsigned)r);
        lq(id ? id : "");
        sb_put(&logb, "\n");
        log_flush();
    }
    return r;
}

lbm_ulong lbm_read_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags, lbm_message **out)
{
    ensure_store();
    lbm_ulong r = LBM_SUCCESS;
    if (out) *out = NULL;
    int n = 0;
    entry *e = NULL;
    msg m = { 0 };
    int i = -1;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    else if (!out) r = LBM_E_FAILURE;
    else {
        e = inbox(&n);
        i = id ? find_entry(e, n, id) : -1;
        if (i < 0 || load_msg(e[i].file, &m)) r = LBM_E_INVALID_MESSAGE;
    }
    if (!r) {
        arena *a = anew();
        lbm_message *o = (lbm_message *)aalloc(a, sizeof *o);
        a->top = o;
        o->lpszSubject = astr(a, m.subject);
        o->lpszMessageType = astr(a, m.type);
        o->lpszDateReceived = astr(a, m.date);
        o->lpszConversationID = astr(a, m.conv);
        o->flFlags = m.flags | (e[i].unread ? LBM_UNREAD : 0);
        o->lpOriginator = (lbm_recip *)aalloc(a, sizeof(lbm_recip));   /* always one, empty without From (as the
                                                                           mail programs: MSMAPI32.OCX fails on NULL) */
        o->lpOriginator->lpszName = astr(a, m.has_orig ? m.orig.name : "");
        o->lpOriginator->lpszAddress = astr(a, m.has_orig ? m.orig.address : "");
        o->nRecipCount = (lbm_ulong)m.nrc;
        o->lpRecips = m.nrc ? (lbm_recip *)aalloc(a, sizeof(lbm_recip) * (size_t)m.nrc) : NULL;
        for (int k = 0; k < m.nrc; k++) {
            o->lpRecips[k].ulRecipClass = m.rc[k].cls;
            o->lpRecips[k].lpszName = astr(a, m.rc[k].name);
            o->lpRecips[k].lpszAddress = astr(a, m.rc[k].address);
        }
        int body_file = (flags & LBM_BODY_AS_FILE) && !(flags & LBM_ENVELOPE_ONLY);
        if (!(flags & LBM_ENVELOPE_ONLY) && !body_file) o->lpszNoteText = astr(a, m.note);
        int nfiles = (flags & (LBM_ENVELOPE_ONLY | LBM_SUPPRESS_ATTACH)) ? 0 : m.nat;
        int total = nfiles + (body_file ? 1 : 0);
        if (total) {
            o->lpFiles = (lbm_file *)aalloc(a, sizeof(lbm_file) * (size_t)total);
            int k = 0;
            if (body_file) {                   /* the text as the first attachment */
                sbuf rel = { 0 };
                sb_printf(&rel, "tmp/%s_body.txt", id);
                char *path = spath(rel.p);
                write_file(path, m.note, strlen(m.note));
                sbuf shown = { 0 };
                sb_printf(&shown, "%s\\tmp\\%s_body.txt", lbm_store(), id);
                o->lpFiles[k].nPosition = 0xFFFFFFFFu;
                o->lpFiles[k].lpszPathName = astr(a, shown.p);
                o->lpFiles[k].lpszFileName = NULL;
                free(path); free(rel.p); free(shown.p);
                k++;
            }
            for (int j = 0; j < nfiles; j++, k++) {
                sbuf rel = { 0 };
                sb_printf(&rel, "tmp/%s_%d_%s", id, j + 1, m.at[j].name);
                char *path = spath(rel.p);
                if (write_file(path, m.at[j].data, m.at[j].len)) r = LBM_E_ATTACHMENT_WRITE_FAILURE;
                sbuf shown = { 0 };            /* the path the program gets: "\" after the store, as on Windows */
                sb_printf(&shown, "%s\\tmp\\%s_%d_%s", lbm_store(), id, j + 1, m.at[j].name);
                o->lpFiles[k].flFlags = m.at[j].flags;
                o->lpFiles[k].nPosition = m.at[j].pos;
                o->lpFiles[k].lpszPathName = astr(a, shown.p);
                o->lpFiles[k].lpszFileName = astr(a, m.at[j].name);
                free(path); free(rel.p); free(shown.p);
            }
        }
        o->nFileCount = (lbm_ulong)total;
        if (r) { lbm_free_buffer(o); o = NULL; }
        else {
            *out = o;
            if (!(flags & LBM_PEEK) && e[i].unread) {   /* read: to cur/ with the flag S */
                sbuf rel = { 0 };
                sb_printf(&rel, "inbox/cur/%s" FLAGSEP "S", e[i].id);
                char *to = spath(rel.p);
                char *from = xstrdup(e[i].file);
                native(from);
                rename(from, to);
                free(to); free(from); free(rel.p);
            }
        }
    }
    if (logging()) {
        sb_printf(&logb, "MAPIReadMail(session=0x%lx, ui=%s, id=", (unsigned long)session, ui_str(ui));
        lq(id);
        sb_printf(&logb, ", flags=0x%x) = %u\n", (unsigned)flags, (unsigned)r);
        if (!r && out && *out) log_message(*out);
        log_flush();
    }
    free_msg(&m);
    free_inbox(e, n);
    return r;
}

lbm_ulong lbm_save_mail(lbm_handle session, lbm_handle ui, const lbm_message *m, lbm_ulong flags, char *id)
{
    ensure_store();
    lbm_ulong r = LBM_SUCCESS;
    int lg = logging();
    if (lg) {
        sb_printf(&logb, "MAPISaveMail(session=0x%lx, ui=%s, flags=0x%x, id=", (unsigned long)session, ui_str(ui), (unsigned)flags);
        lq(id);
        sb_put(&logb, ")\n");
        log_message(m);
    }
    char *newid = NULL;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    else if (!m) r = LBM_E_FAILURE;
    else {
        int n;
        entry *e = inbox(&n);
        if (id && *id) {                       /* the message replaced: its file goes, its id stays */
            int i = find_entry(e, n, id);
            if (i < 0) r = LBM_E_INVALID_MESSAGE;
            else { char *f = xstrdup(e[i].file); native(f); remove(f); free(f); newid = xstrdup(id); }
        } else newid = next_id();
        free_inbox(e, n);
        if (!r) {
            sbuf o = { 0 };
            r = compose_file(&o, m, NULL, 0);
            if (!r) {
                sbuf rel = { 0 };
                if (m->flFlags & LBM_UNREAD) sb_printf(&rel, "inbox/new/%s", newid);
                else sb_printf(&rel, "inbox/cur/%s" FLAGSEP "S", newid);
                char *path = spath(rel.p);
                if (write_file(path, o.p ? o.p : "", o.n)) r = LBM_E_DISK_FULL;
                free(path); free(rel.p);
            }
            free(o.p);
        }
    }
    if (!r && id) snprintf(id, 512, "%s", newid);
    if (lg) { sb_printf(&logb, "  = %u id=", (unsigned)r); lq(r ? "" : newid); sb_put(&logb, "\n"); log_flush(); }
    free(newid);
    return r;
}

lbm_ulong lbm_delete_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags)
{
    ensure_store();
    lbm_ulong r = LBM_SUCCESS;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    else {
        int n;
        entry *e = inbox(&n);
        int i = id ? find_entry(e, n, id) : -1;
        if (i < 0) r = LBM_E_INVALID_MESSAGE;
        else { char *f = xstrdup(e[i].file); native(f); if (remove(f)) r = LBM_E_ACCESS_DENIED; free(f); }
        free_inbox(e, n);
    }
    if (logging()) {
        sb_printf(&logb, "MAPIDeleteMail(session=0x%lx, ui=%s, id=", (unsigned long)session, ui_str(ui));
        lq(id);
        sb_printf(&logb, ", flags=0x%x) = %u\n", (unsigned)flags, (unsigned)r);
        log_flush();
    }
    return r;
}

lbm_ulong lbm_address(lbm_handle session, lbm_handle ui, const char *caption, lbm_ulong nedit, const char *labels,
                      lbm_ulong nrecips, const lbm_recip *recips, lbm_ulong flags, lbm_ulong *nnew, lbm_recip **news)
{
    lbm_ulong r = LBM_SUCCESS;
    int lg = logging();
    if (lg) {
        sb_printf(&logb, "MAPIAddress(session=0x%lx, ui=%s, caption=", (unsigned long)session, ui_str(ui));
        lq(caption);
        sb_printf(&logb, ", editfields=%u, labels=", (unsigned)nedit);
        lq(labels);
        sb_printf(&logb, ", recips=%u, flags=0x%x)\n", (unsigned)nrecips, (unsigned)flags);
        for (lbm_ulong i = 0; i < nrecips && recips; i++) {
            char w[32];
            snprintf(w, sizeof w, "recip(%u)", (unsigned)i);
            log_recip(w, &recips[i]);
        }
    }
    if (nnew) *nnew = 0;
    if (news) *news = NULL;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    else if (nedit > 4) r = LBM_E_INVALID_EDITFIELDS;
    else if (!dialog_sends()) r = LBM_USER_ABORT;
    else if (nnew && news) {                   /* OK: the recipients it was given, resolved */
        arena *a = anew();
        lbm_recip *o = (lbm_recip *)aalloc(a, sizeof(lbm_recip) * (nrecips ? nrecips : 1));
        a->top = o;
        for (lbm_ulong i = 0; i < nrecips; i++) {
            o[i].ulRecipClass = recips[i].ulRecipClass ? recips[i].ulRecipClass : 1;
            char *n = NULL, *ad = NULL;
            if ((!recips[i].lpszAddress || !*recips[i].lpszAddress) && !resolve(recips[i].lpszName, 1, &n, &ad)) {
                o[i].lpszName = astr(a, n);
                o[i].lpszAddress = astr(a, ad);
            } else {
                o[i].lpszName = astr(a, recips[i].lpszName ? recips[i].lpszName : "");
                o[i].lpszAddress = astr(a, recips[i].lpszAddress ? recips[i].lpszAddress : "");
            }
            free(n); free(ad);
        }
        *nnew = nrecips;
        *news = o;
        if (!nrecips) { *news = NULL; lbm_free_buffer(o); }
    }
    if (lg) {
        sb_printf(&logb, "  = %u new=%u\n", (unsigned)r, (unsigned)(nnew ? *nnew : 0));
        for (lbm_ulong i = 0; !r && nnew && news && *news && i < *nnew; i++) {
            char w[32];
            snprintf(w, sizeof w, "new(%u)", (unsigned)i);
            log_recip(w, &(*news)[i]);
        }
        log_flush();
    }
    return r;
}

lbm_ulong lbm_details(lbm_handle session, lbm_handle ui, const lbm_recip *recip, lbm_ulong flags)
{
    lbm_ulong r = LBM_SUCCESS;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    else if (!recip) r = LBM_E_FAILURE;
    else if (!dialog_sends()) r = LBM_USER_ABORT;
    if (logging()) {
        sb_printf(&logb, "MAPIDetails(session=0x%lx, ui=%s, flags=0x%x) = %u\n", (unsigned long)session, ui_str(ui),
                  (unsigned)flags, (unsigned)r);
        log_recip("recip", recip);
        log_flush();
    }
    return r;
}

lbm_ulong lbm_resolve_name(lbm_handle session, lbm_handle ui, const char *name, lbm_ulong flags, lbm_recip **out)
{
    lbm_ulong r = LBM_SUCCESS;
    char *n = NULL, *a = NULL;
    if (out) *out = NULL;
    if (!session_ok(session)) r = LBM_E_INVALID_SESSION;
    else if (!out) r = LBM_E_FAILURE;
    else r = resolve(name, (flags & LBM_DIALOG) != 0, &n, &a);
    if (!r) {
        arena *ar = anew();
        lbm_recip *o = (lbm_recip *)aalloc(ar, sizeof(lbm_recip));
        ar->top = o;
        o->ulRecipClass = 1;
        o->lpszName = astr(ar, n);
        o->lpszAddress = astr(ar, a);
        *out = o;
    }
    if (logging()) {
        sb_printf(&logb, "MAPIResolveName(session=0x%lx, ui=%s, name=", (unsigned long)session, ui_str(ui));
        lq(name);
        sb_printf(&logb, ", flags=0x%x) = %u\n", (unsigned)flags, (unsigned)r);
        if (!r) log_recip("recip", *out);
        log_flush();
    }
    free(n); free(a);
    return r;
}
