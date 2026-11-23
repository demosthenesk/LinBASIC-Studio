#define _GNU_SOURCE
/* io.c - Debug.Print, Print #, Write #, files and paths.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <math.h>
#include "vbrt_int.h"
#include "vbos.h"
#include <errno.h>
#include <strings.h>
#include <ctype.h>
#include <time.h>

#define NFILES 1024                             /* (VB6 takes #512 too) */
static vbfile files[NFILES];
static char *fpath[NFILES];                     /* the full path of each open file (Open, Kill and Name check it) */
static int debug_col;

/* VB paths use '\'; on Linux they become '/' */
static char *native_path(vbs p)
{
    char *s = vb_stoutf8(p);
#ifndef _WIN32
    for (char *q = s; *q; q++) if (*q == '\\') *q = '/';
    vbos_ci_path(s);                            /* "Jumble.skn" opens JUMBLE.SKN, as on Windows */
#endif
    return s;
}

static vbfile *getf(vb_i32 n, int need_mode)
{
    if (n < 1 || n >= NFILES) { vb_raise(52); return 0; }
    vbfile *f = &files[n];
    if (!f->open) { vb_raise(52); return 0; }
    if (need_mode && !(f->mode & need_mode)) { vb_raise(54); return 0; }
    return f;
}

vbfile *vb_getfile(vb_i32 n, int need_mode) { return getf(n, need_mode); }

void vb_files_flush(void) { for (int i = 1; i < NFILES; i++) if (files[i].open) fflush(files[i].f); }

/* the full path of a file that exists (0 if it does not) */
static char *full_path(const char *p)
{
#ifdef _WIN32
    char b[1024];
    if (!_fullpath(b, p, sizeof b)) return 0;
    for (char *q = b; *q; q++) *q = (char)tolower((unsigned char)*q);
    return strdup(b);
#else
    return realpath(p, 0);
#endif
}
/* the number of the file open with that path (0: none) */
static int open_as(const char *p)
{
    char *fp = full_path(p);
    if (!fp) return 0;
    int r = 0;
    for (int i = 1; i < NFILES && !r; i++) if (files[i].open && fpath[i] && !strcmp(fpath[i], fp)) r = i;
    free(fp);
    return r;
}

void vb_open(vbs path, vb_i32 mode, vb_i32 fileno, vb_i32 reclen)
{
    if (fileno < 1 || fileno >= NFILES) { vb_raise(52); return; }
    if (files[fileno].open) { vb_raise(55); return; }
    char *p = native_path(path);
    if (!*p) { vb_raise(75); return; }
    {   /* open already: only Input, Binary or Random again in the same mode (VB6) */
        int k = open_as(p);
        if (k && (mode == VBF_OUTPUT || mode == VBF_APPEND || files[k].mode != mode)) { vb_raise(55); return; }
    }
    const char *m = mode == VBF_INPUT ? "rb" : mode == VBF_OUTPUT ? "wb" : mode == VBF_APPEND ? "ab" : "r+b";
    FILE *f = vbos_fopen(p, m);
    if (!f && (mode == VBF_BINARY || mode == VBF_RANDOM)) f = vbos_fopen(p, "w+b");
    if (!f) { vb_raise(errno == ENOENT ? (mode == VBF_INPUT ? 53 : 76) : errno == EACCES ? 75 : 75); return; }
    if (getenv("VBRT_UNBUFFERED")) setvbuf(f, 0, _IONBF, 0);   /* (tests: what a program that hangs wrote so far) */
    files[fileno].f = f; files[fileno].mode = mode; files[fileno].col = 0;
    files[fileno].reclen = reclen > 0 ? reclen : 128; files[fileno].open = 1;
    files[fileno].eofflag = 0; files[fileno].lastrec = 0;
    free(fpath[fileno]); fpath[fileno] = full_path(p);
}

void vb_close(vb_i32 fileno)
{
    if (fileno == 0) { for (int i = 1; i < NFILES; i++) if (files[i].open) { fclose(files[i].f); files[i].open = 0; free(fpath[i]); fpath[i] = 0; } return; }
    /* VB6: Close of a file number that is not open does nothing (no error 52) */
    if (fileno < 1 || fileno >= NFILES) { vb_raise(52); return; }
    vbfile *f = &files[fileno];
    if (!f->open) return;
    fclose(f->f); f->open = 0;
    free(fpath[fileno]); fpath[fileno] = 0;
}

vb_i32 vb_freefile(void) { for (int i = 1; i < 256; i++) if (!files[i].open) return i; vb_raise(67); return 0; }

/* ---------------------------------------------------------------- printing */
static int pr_ch;
static vbfile *pr_f;
/* pr_ch == -2: Print method of a form or PictureBox: the text goes to a buffer (UTF-8,
 * a TAB for each comma) that the forms library draws */
static char *pb_buf; static int pb_len, pb_cap, pb_col, pb_nl;
static void pb_add(const char *s, int n)
{
    if (pb_len + n + 1 > pb_cap) { pb_cap = (pb_len + n + 1) * 2; pb_buf = (char *)realloc(pb_buf, (size_t)pb_cap); }
    memcpy(pb_buf + pb_len, s, (size_t)n); pb_len += n; pb_buf[pb_len] = 0;
    for (int i = 0; i < n; i++) if ((s[i] & 0xC0) != 0x80) pb_col++;
}

static void ch_write(const char *s, int n)
{
    if (pr_ch == -2) pb_add(s, n);
    else if (pr_ch == 0) {
        /* Debug window: UTF-8 on standard output */
        fwrite(s, 1, n, stdout);
        for (int i = 0; i < n; i++) if ((s[i] & 0xC0) != 0x80) debug_col++;
    } else if (pr_f) {
        fwrite(s, 1, n, pr_f->f);
        pr_f->col += n;
    }
}
static void ch_nl(void)
{
    if (pr_ch == -2) { pb_add("\n", 1); pb_col = 0; }
    else if (pr_ch == 0) { fputs("\n", stdout); debug_col = 0; }
    else if (pr_f) { fputs("\r\n", pr_f->f); pr_f->col = 0; }
}
static int ch_col(void) { return pr_ch == -2 ? pb_col : pr_ch == 0 ? debug_col : pr_f ? pr_f->col : 0; }

static void put_vbs(vbs s)
{
    if (pr_ch <= 0) { char *u = vb_stoutf8(s); ch_write(u, (int)strlen(u)); }
    else { int n; char *a = vb_stoansi(s, &n); ch_write(a, n); }
}

/* a Print inside the expression of a Print (Print #f, "pk " & uc.MoveToPK(3), whose events Print #f too:
 * VB6 oracle do1): each Print keeps the state of the one around it and gives it back at its end */
static int wr_count;
typedef struct prstate { int ch; vbfile *f; int wr; char *buf; int len, cap, col, nl; } prstate;
#define PR_DEPTH 64
static prstate pr_stack[PR_DEPTH]; static int pr_depth;
static char *pb_done; static int pb_done_nl;   /* the text of the Print method that ended (vb_pr_take) */
static void pr_push(void)
{
    if (pr_depth == PR_DEPTH) {                  /* Prints left by errors: the oldest go */
        if (pr_stack[0].buf) free(pr_stack[0].buf);
        memmove(pr_stack, pr_stack + 1, sizeof(prstate) * (PR_DEPTH - 1));
        pr_depth--;
    }
    prstate *x = &pr_stack[pr_depth++];
    x->ch = pr_ch; x->f = pr_f; x->wr = wr_count;
    x->buf = pb_buf; x->len = pb_len; x->cap = pb_cap; x->col = pb_col; x->nl = pb_nl;
    pb_buf = 0; pb_len = pb_cap = pb_col = pb_nl = 0;
}
static void pr_pop(void)
{
    free(pb_buf);
    if (!pr_depth) { pr_ch = 0; pr_f = 0; wr_count = 0; pb_buf = 0; pb_len = pb_cap = pb_col = pb_nl = 0; return; }
    prstate *x = &pr_stack[--pr_depth];
    pr_ch = x->ch; pr_f = x->f; wr_count = x->wr;
    pb_buf = x->buf; pb_len = x->len; pb_cap = x->cap; pb_col = x->col; pb_nl = x->nl;
    if (pr_f && !pr_f->open) { pr_f = 0; pr_ch = -1; }   /* its file was closed meanwhile */
}

void vb_pr_begin(vb_i32 fileno)
{
    pr_push();
    pr_ch = fileno; pr_f = 0;
    if (fileno == -2) { pb_add("", 0); return; }
    if (fileno) { pr_f = getf(fileno, VBF_OUTPUT | VBF_APPEND); if (!pr_f) pr_ch = -1; }
}
static int live(void) { return !vb_errpend && (pr_ch >= 0 || pr_ch == -2); }

void vb_pr_s(vbs s) { if (live()) put_vbs(s); }
static void pr_num(const char *txt) { char b[64]; snprintf(b, sizeof b, "%s%s ", txt[0] == '-' ? "" : " ", txt); ch_write(b, (int)strlen(b)); }
void vb_pr_i(vb_i32 v) { if (!live()) return; char b[24]; snprintf(b, sizeof b, "%d", v); pr_num(b); }
void vb_pr_f(float v) { if (!live()) return; char b[48]; vb_fmt_real(v, 7, VB_FMT_LOCALE, b); pr_num(b); }
void vb_pr_d(double v) { if (!live()) return; char b[48]; vb_fmt_real(v, 15, VB_FMT_LOCALE, b); pr_num(b); }
void vb_pr_c(vb_cy v) { if (!live()) return; char b[48]; vb_fmt_cy(v, VB_FMT_LOCALE, b); pr_num(b); }
void vb_pr_bool(vb_bool v) { if (live()) ch_write(v ? "True" : "False", v ? 4 : 5); }
void vb_pr_date(vb_date v) { if (!live()) return; vb_i32 m = vb_mark(); put_vbs(vb_cstr_date(v)); ch_write(" ", 1); vb_release(m); }   /* VB6: a blank after a date, none before */
void vb_pr_v(vbvar v)
{
    if (!live()) return;
    switch (v.vt) {
    case VT_EMPTY: break;
    case VT_NULL: ch_write("Null", 4); break;
    case VT_I2: vb_pr_i(v.u.i); break;
    case VT_I4: vb_pr_i(v.u.l); break;
    case VT_UI1: vb_pr_i(v.u.b); break;
    case VT_R4: vb_pr_f(v.u.f); break;
    case VT_R8: vb_pr_d(v.u.d); break;
    case VT_CY: vb_pr_c(v.u.cy); break;
    case VT_DECIMAL: { char b[100]; vb_dec_text(v, 1, b); pr_num(b); break; }
    case VT_BOOL: vb_pr_bool(v.u.i); break;
    case VT_DATE: vb_pr_date(v.u.d); break;
    case VT_BSTR: put_vbs(v.u.s); break;
    case VT_ERROR: { char b[32]; snprintf(b, sizeof b, "Error %d", v.u.l & 0xFFFF); ch_write(b, (int)strlen(b)); break; }
    case VT_DISPATCH: { vbvar d = vb_late(v, "", VBM_GET, 0, 0); if (!vb_errpend) vb_pr_v(d); break; }   /* default member */
    }
}
void vb_pr_comma(void)
{
    if (!live()) return;
    if (pr_ch == -2) { pb_add("\t", 1); return; }
    int c = ch_col(), next = (c / 14 + 1) * 14;
    while (ch_col() < next) ch_write(" ", 1);
}
void vb_pr_tab(vb_i32 col)
{
    if (!live()) return;
    if (col <= 0) { vb_pr_comma(); return; }
    int target = col - 1;
    if (ch_col() > target) ch_nl();
    while (ch_col() < target) ch_write(" ", 1);
}
void vb_pr_spc(vb_i32 n) { if (!live()) return; for (int i = 0; i < n; i++) ch_write(" ", 1); }
vbs vb_pr_take(vb_i32 *newline)
{
    *newline = pb_done_nl;
    return vb_sfromutf8(pb_done ? pb_done : "");
}
void vb_pr_end(vb_i32 newline)
{
    if (pr_ch == -2) {
        free(pb_done); pb_done = pb_buf; pb_done_nl = newline;
        pb_buf = 0; pb_len = pb_cap = 0;
        pr_pop();
        return;
    }
    if (live() && newline) ch_nl();
    if (pr_ch == 0) fflush(stdout);
    pr_pop();
}

/* Write #: strings in quotes, numbers without spaces, #TRUE#, #NULL#, dates as #yyyy-mm-dd# */
void vb_wr_begin(vb_i32 fileno) { vb_pr_begin(fileno); wr_count = 0; }
void vb_wr_v(vbvar v)
{
    if (!live()) return;
    if (wr_count++) ch_write(",", 1);
    char b[64];
    switch (v.vt) {
    case VT_EMPTY: break;
    case VT_NULL: ch_write("#NULL#", 6); break;
    case VT_BOOL: ch_write(v.u.i ? "#TRUE#" : "#FALSE#", v.u.i ? 6 : 7); break;
    case VT_BSTR: ch_write("\"", 1); put_vbs(v.u.s); ch_write("\"", 1); break;
    case VT_DATE: vb_date_iso(v.u.d, b); ch_write(b, (int)strlen(b)); break;
    /* Write # is independent of the regional settings */
    case VT_R4: vb_fmt_real(v.u.f, 7, VB_FMT_STR, b); ch_write(b, (int)strlen(b)); break;
    case VT_R8: vb_fmt_real(v.u.d, 15, VB_FMT_STR, b); ch_write(b, (int)strlen(b)); break;
    case VT_CY: vb_fmt_cy(v.u.cy, VB_FMT_STR, b); ch_write(b, (int)strlen(b)); break;
    case VT_DECIMAL: { char t[100]; vb_dec_text(v, 0, t); ch_write(t, (int)strlen(t)); break; }
    default: { vb_i32 m = vb_mark(); put_vbs(vb_v2s(v)); vb_release(m); }
    }
}
void vb_wr_end(void) { if (live()) ch_nl(); pr_pop(); }

/* ---------------------------------------------------------------- reading */
static int rdc(vbfile *f) { return fgetc(f->f); }

void vb_line_input(vb_i32 fileno, vbs *dst)
{
    vbfile *f = getf(fileno, VBF_INPUT);
    if (!f) return;
    int c = rdc(f);
    if (c == EOF) { vb_raise(62); return; }
    char *buf = 0; size_t n = 0, cap = 0;
    while (c != EOF && c != '\n' && c != '\r') {
        if (n + 1 >= cap) { cap = cap ? cap * 2 : 128; buf = (char *)realloc(buf, cap); }
        buf[n++] = (char)c; c = rdc(f);
    }
    if (c == '\r') { int d = rdc(f); if (d != '\n' && d != EOF) ungetc(d, f->f); }
    vb_i32 m = vb_mark();
    vb_sset(dst, vb_sfromansi(buf ? buf : "", (int)n));
    vb_release(m);
    free(buf);
}

/* one field of Input #: quoted string, or text up to a comma / end of line */
static vbs read_field(vbfile *f, int *quoted)
{
    int c;
    do c = rdc(f); while (c == ' ' || c == '\t');
    if (c == EOF) { vb_raise(62); return vb_stmp(0); }
    char *buf = 0; size_t n = 0, cap = 0;
#define PUT(ch) do { if (n + 1 >= cap) { cap = cap ? cap * 2 : 64; buf = (char *)realloc(buf, cap); } buf[n++] = (char)(ch); } while (0)
    *quoted = 0;
    if (c == '"') {
        *quoted = 1;
        while ((c = rdc(f)) != EOF && c != '"') PUT(c);
        do c = rdc(f); while (c == ' ' || c == '\t');
    } else {
        while (c != EOF && c != ',' && c != '\n' && c != '\r') { PUT(c); c = rdc(f); }
        while (n && (buf[n - 1] == ' ' || buf[n - 1] == '\t')) n--;
    }
    if (c == '\r') { int d = rdc(f); if (d != '\n' && d != EOF) ungetc(d, f->f); }
#undef PUT
    vbs r = vb_sfromansi(buf ? buf : "", (int)n);
    free(buf);
    return r;
}

void vb_input_s(vb_i32 fileno, vbs *dst)
{
    vbfile *f = getf(fileno, VBF_INPUT); if (!f) return;
    int q; vb_i32 m = vb_mark(); vbs s = read_field(f, &q);
    if (!vb_errpend) vb_sset(dst, s);
    vb_release(m);
}
void vb_input_d(vb_i32 fileno, double *dst)
{
    vbfile *f = getf(fileno, VBF_INPUT); if (!f) return;
    int q; vb_i32 m = vb_mark(); vbs s = read_field(f, &q);
    if (!vb_errpend) {
        vb_i32 m2 = vb_mark(); vbs t = vb_ucase(s);
        if (t.len == 6 && t.p[0] == '#') *dst = (t.p[1] == 'T') ? -1 : 0;   /* #TRUE# / #FALSE# */
        else *dst = vb_val(s);
        vb_release(m2);
    }
    vb_release(m);
}
/* a field of Input # that is a number as Write # / Print # writes it: [-+]digits[.digits][E[-+]digits] */
/* a number as Write # writes it; *frac = its digits after the point (-1 with an exponent) */
static int plain_number(vbs s, int *frac)
{
    vb_i32 i = 0, n = s.len, digits = 0;
    *frac = 0;
    while (i < n && s.p[i] == ' ') i++;
    if (i < n && (s.p[i] == '-' || s.p[i] == '+')) i++;
    while (i < n && s.p[i] >= '0' && s.p[i] <= '9') { i++; digits++; }
    if (i < n && s.p[i] == '.') { i++; while (i < n && s.p[i] >= '0' && s.p[i] <= '9') { i++; digits++; (*frac)++; } }
    if (!digits) return 0;
    if (i < n && (s.p[i] == 'E' || s.p[i] == 'e' || s.p[i] == 'D' || s.p[i] == 'd')) {
        *frac = -1;
        i++;
        if (i < n && (s.p[i] == '-' || s.p[i] == '+')) i++;
        if (i >= n || s.p[i] < '0' || s.p[i] > '9') return 0;
        while (i < n && s.p[i] >= '0' && s.p[i] <= '9') i++;
    }
    while (i < n && s.p[i] == ' ') i++;
    return i == n;
}
void vb_input_v(vb_i32 fileno, vbvar *dst)
{
    vbfile *f = getf(fileno, VBF_INPUT); if (!f) return;
    int q, frac; vb_i32 m = vb_mark(); vbs s = read_field(f, &q);
    if (!vb_errpend) {
        vbvar v;
        if (q) v = vb_v_s(s);
        else if (s.len == 0) v = vb_v_empty();
        else if (s.len >= 2 && s.p[0] == '#' && s.p[s.len - 1] == '#') {
            vbs u = vb_ucase(s);
            if (u.len == 6 && u.p[1] == 'T') v = vb_v_bool(-1);
            else if (u.len == 7 && u.p[1] == 'F') v = vb_v_bool(0);
            else if (u.len == 6 && u.p[1] == 'N') v = vb_v_null();
            else {
                int y = 0, mo = 0, d = 0;
                char *a = vb_stoutf8(s);
                if (sscanf(a, "#%d-%d-%d#", &y, &mo, &d) == 3) v = vb_v_date(vb_dateserial(y, mo, d));
                else v = vb_v_s(s);
            }
        }
        else if (!plain_number(s, &frac)) v = vb_v_s(s);
        else {                                  /* a number of Write #: "." whatever the locale (Val) */
            double d = vb_val(s);
            /* VB6 (t_ph68): Integer, Long, then Currency for up to 4 decimals ("3.5"), else Double */
            v = (d == (vb_i16)d) ? vb_v_i((vb_i16)d) : (d == (vb_i32)d) ? vb_v_l((vb_i32)d)
              : (frac >= 0 && frac <= 4 && fabs(d) < 922337203685477.0) ? vb_v_c((vb_cy)llround(d * 10000.0))
              : vb_v_d(d);
        }
        vb_vset(dst, v);
    }
    vb_release(m);
}

vb_bool vb_eof(vb_i32 fileno)
{
    vbfile *f = getf(fileno, 0); if (!f) return 0;
    if (f->mode & (VBF_BINARY | VBF_RANDOM)) return f->eofflag ? -1 : 0;   /* the last Get went past the end */
    int c = fgetc(f->f);
    if (c == EOF) return -1;
    ungetc(c, f->f);
    return 0;
}
/* FileAttr(n, 1): the mode it was opened in (Input 1, Output 2, Random 4, Append 8, Binary 32); 2: the file handle */
vb_i32 vb_fileattr(vb_i32 n, vb_i32 type)
{
    vbfile *f = getf(n, 0); if (!f) return 0;
    if (type == 1) return f->mode;
    if (type == 2) return (vb_i32)fileno(f->f);
    vb_raise(5);
    return 0;
}
vb_i32 vb_lof(vb_i32 fileno)
{
    vbfile *f = getf(fileno, 0); if (!f) return 0;
    long pos = ftell(f->f); fseek(f->f, 0, SEEK_END); long n = ftell(f->f); fseek(f->f, pos, SEEK_SET);
    return (vb_i32)n;
}

/* ---------------------------------------------------------------- file system */
void vb_kill(vbs path)
{
    char *p = native_path(path);
    char *slash = strrchr(p, '/');
#ifdef _WIN32
    char *bs = strrchr(p, '\\'); if (bs > slash) slash = bs;
#endif
    const char *name = slash ? slash + 1 : p;
    if (strpbrk(name, "*?")) {                  /* Kill "*.tmp": every file that matches (53 if none) */
        char dir[4096], pat[1024], full[5200];
        if (slash) { snprintf(dir, sizeof dir, "%.*s", (int)(slash - p), p); if (!dir[0]) snprintf(dir, sizeof dir, "/"); }
        else snprintf(dir, sizeof dir, ".");
        snprintf(pat, sizeof pat, "%s", name);
        vbos_dir *d = vbos_opendir(dir);
        if (!d) { vb_raise(76); return; }
        int found = 0, err = 0, isdir;
        const char *e;
        while ((e = vbos_readdir(d, &isdir)) != 0) {
            if (isdir || !vbos_wildmatch(pat, e)) continue;
            snprintf(full, sizeof full, "%s/%s", dir, e);
            found++;
            if (open_as(full)) { err = 55; continue; }
            if (vbos_remove(full) != 0) err = 75;
        }
        vbos_closedir(d);
        if (!found) vb_raise(53); else if (err) vb_raise(err);
        return;
    }
    if (open_as(p)) { vb_raise(55); return; }
    if (vbos_remove(p) != 0) vb_raise(errno == ENOENT ? 53 : 75);
}
vb_i32 vb_filelen(vbs path) { long long n; char *p = native_path(path); if (vbos_stat(p, &n, 0, 0) != 0) { vb_raise(53); return 0; } return (vb_i32)n; }
void vb_mkdir(vbs path) { if (vbos_mkdir(native_path(path)) != 0) vb_raise(errno == ENOENT ? 76 : 75); }
void vb_rmdir(vbs path) { if (vbos_rmdir(native_path(path)) != 0) vb_raise(errno == ENOENT ? 76 : 75); }
void vb_chdir(vbs path) { if (vbos_chdir(native_path(path)) != 0) vb_raise(76); }
vbs vb_curdir(void) { char b[4096]; vbos_getcwd(b, sizeof b); return vb_sfromutf8(b); }
/* CurDir(drive): a drive is a letter (error 5 otherwise); here every drive is the current folder */
vbs vb_curdir_d(vbs d)
{
    if (d.len && !((d.p[0] >= 'A' && d.p[0] <= 'Z') || (d.p[0] >= 'a' && d.p[0] <= 'z'))) { vb_raise(5); return vb_stmp(0); }
    return vb_curdir();
}
void vb_name(vbs from, vbs to) { char *a = native_path(from); if (open_as(a)) { vb_raise(55); return; } char *b = native_path(to); if (vbos_rename(a, b) != 0) vb_raise(errno == ENOENT ? 53 : 75); }

/* FileCopy, FileDateTime, GetAttr */
void vb_filecopy(vbs from, vbs to)
{
    char *a = strdup(native_path(from)), *b = strdup(native_path(to));
    int r = vbos_copyfile(a, b);
    if (r == -1) vb_raise(53); else if (r) vb_raise(75);
    free(a); free(b);
}
vb_date vb_filedatetime(vbs path)
{
    long long t;
    if (vbos_stat(native_path(path), 0, &t, 0) != 0) { vb_raise(53); return 0; }
    time_t tt = (time_t)t;
    struct tm lt;
#ifdef _WIN32
    localtime_s(&lt, &tt);
#else
    localtime_r(&tt, &lt);
#endif
    return vb_dateserial(lt.tm_year + 1900, lt.tm_mon + 1, lt.tm_mday) +
           (lt.tm_hour * 3600 + lt.tm_min * 60 + lt.tm_sec) / 86400.0;
}
void vb_setattr(vbs path, vb_i32 attr)
{
    char *p = native_path(path);
    int a;
    if (vbos_stat(p, 0, 0, &a) != 0) { vb_raise(53); return; }
    if (vbos_setattr(p, attr) != 0) vb_raise(75);
}
void vb_chdrive(vbs drive)
{
#ifdef _WIN32
    if (drive.len && vbos_chdir(native_path(vb_scat(vb_left(drive, 1), vb_sfromutf8(":")))) != 0) vb_raise(68);
#else
    (void)drive;                                        /* one tree of folders: no drives */
#endif
}
vb_i32 vb_getattr(vbs path)
{
    int a;
    if (vbos_stat(native_path(path), 0, 0, &a) != 0) { vb_raise(53); return 0; }
    return a;
}

static char dir_path[4096], dir_pat[1024];
static vbos_dir *dir_h;
static vb_i32 dir_attr;
vbs vb_dir(vbs pattern, vb_i32 attr, vb_i32 has_pattern)
{
    if (has_pattern) {
        char *p = native_path(pattern);
        char *slash = strrchr(p, '/'), *bs = strrchr(p, '\\');
        if (bs > slash) slash = bs;
        if (slash) { size_t k = (size_t)(slash - p); memcpy(dir_path, p, k); dir_path[k] = 0; snprintf(dir_pat, sizeof dir_pat, "%s", slash + 1); }
        else { strcpy(dir_path, "."); snprintf(dir_pat, sizeof dir_pat, "%s", p); }
        if (slash && slash == p) strcpy(dir_path, "/");
        if (!dir_pat[0]) strcpy(dir_pat, "*");
        if (!strcmp(dir_pat, "*.*")) strcpy(dir_pat, "*");
        if (dir_h) vbos_closedir(dir_h);
        dir_h = vbos_opendir(dir_path);
        dir_attr = attr;
    } else if (!dir_h) { vb_raise(5); return vb_stmp(0); }
    const char *e;
    int isdir;
    while (dir_h && (e = vbos_readdir(dir_h, &isdir))) {
        if (isdir && !(dir_attr & 16)) continue;           /* folders only with vbDirectory */
        if (vbos_wildmatch(dir_pat, e)) return vb_sfromutf8(e);
    }
    if (dir_h) { vbos_closedir(dir_h); dir_h = 0; }
    return vb_stmp(0);
}

vbs vb_app_path(void)
{
    extern const char *vb_argv0(void);
    char b[4096];
    long n = -1;
    const char *ide = getenv("VBRT_APP_PATH");            /* run from the IDE: App.Path is the project folder */
    if (ide && *ide) {
        /* without a separator at the end, as VB6's (only a root keeps it): Dir(App.Path) of ".../x/" gave the
         * first file of the folder (robot 2026-11-12: Code Statistics, a folder taken for a file) */
        snprintf(b, sizeof b, "%s", ide);
        n = (long)strlen(b);
        while (n > 1 && (b[n - 1] == '/' || b[n - 1] == '\\') && !(n == 3 && b[1] == ':')) b[--n] = 0;
        return vb_sfromutf8(b);
    }
    vbos_exe_path(b, sizeof b);
    n = (long)strlen(b);
    if (n <= 0) { snprintf(b, sizeof b, "%s", vb_argv0()); n = (long)strlen(b); }
    b[n] = 0;
    char *s = strrchr(b, '/'), *bs = strrchr(b, '\\');
    if (bs > s) s = bs;
    if (s) *s = 0; else strcpy(b, ".");
    return vb_sfromutf8(b);
}
vbs vb_app_exename(void)
{
    extern const char *vb_argv0(void);
    const char *a = vb_argv0(), *s = strrchr(a, '/'), *bs = strrchr(a, '\\');
    if (bs > s) s = bs;
    char b[1024]; snprintf(b, sizeof b, "%s", s ? s + 1 : a);
    char *dot = strrchr(b, '.'); if (dot && !strcasecmp(dot, ".exe")) *dot = 0;
    return vb_sfromutf8(b);
}

/* Input(n, #f) / InputB(n, #f): n characters (bytes) of a file opened For Input or Binary */
static char *read_n(vb_i32 n, vb_i32 fileno)
{
    vbfile *f = getf(fileno, VBF_INPUT | VBF_BINARY);
    if (!f) return 0;
    if (n < 0) { vb_raise(5); return 0; }
    char *buf = (char *)malloc(n + 1);
    size_t got = fread(buf, 1, (size_t)n, f->f);
    if ((vb_i32)got < n) { free(buf); vb_raise(62); return 0; }
    return buf;
}

vbs vb_input_n(vb_i32 n, vb_i32 fileno)
{
    char *b = read_n(n, fileno);
    if (!b) return vb_stmp(0);
    vbs r = vb_sfromansi(b, n);
    free(b);
    return r;
}

vbs vb_inputb(vb_i32 n, vb_i32 fileno)
{
    char *b = read_n(n, fileno);
    if (!b) return vb_stmp(0);
    vbs r = vb_sfrombytes(b, n);
    free(b);
    return r;
}

/* Shell, Environ */
double vb_shell(vbs cmd, vb_i32 style)
{
    long pid = vbos_shell(vb_stoutf8(cmd), style);
    if (pid < 0) { vb_raise(53); return 0; }
    return (double)pid;
}
vbs vb_environ_s(vbs name)
{
    const char *v = vbos_getenv(vb_stoutf8(name));
    return v ? vb_sfromutf8(v) : vb_stmp(0);
}
vbs vb_environ_n(vb_i32 n)
{
    if (n < 1) { vb_raise(5); return vb_stmp(0); }
    const char *v = vbos_environ_n(n);
    return v ? vb_sfromutf8(v) : vb_stmp(0);
}
