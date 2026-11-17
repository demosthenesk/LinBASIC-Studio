/* scripting.c - the Microsoft Scripting Runtime (scrrun.dll), written again for LinBASIC Studio:
 * FileSystemObject, Drive(s), Folder(s), File(s), TextStream and Dictionary, with the members,
 * defaults and run-time errors of the original. CreateObject("Scripting.FileSystemObject") and
 * Dim x As New Scripting.Dictionary both give these objects.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include "vbos.h"
#include <errno.h>
#include <time.h>
#include <strings.h>

#ifdef _WIN32
#define SEP '\\'
#define SEPS "\\"
#else
#define SEP '/'
#define SEPS "/"
#endif

extern const vb_class vb_cls_fso, vb_cls_drive, vb_cls_drives, vb_cls_folder, vb_cls_folders, vb_cls_file,
    vb_cls_files, vb_cls_textstream, vb_cls_dictionary;

/* ---------------------------------------------------------------- helpers */
static int is_sep(char c) { return c == '/' || c == '\\'; }
static int missing(vb_i32 argc, vbvar *argv, int i) { return i >= argc || vb_ismissing(argv[i]); }
static vbvar vstr(const char *s) { return vb_v_s(vb_sfromutf8(s)); }

/* a path argument: UTF-8, with '/' on Linux (VB programs write '\') */
static char *apath(vbvar v)
{
    vbs s = vb_v2s(v);
    if (vb_errpend) return (char *)"";
    char *p = vb_stoutf8(s);
#ifndef _WIN32
    for (char *q = p; *q; q++) if (*q == '\\') *q = '/';
    vbos_ci_path(p);                            /* names without case, as on Windows */
#endif
    return p;
}
static char *astr(vbvar v) { vbs s = vb_v2s(v); return vb_errpend ? (char *)"" : vb_stoutf8(s); }
static vb_bool abool(vb_i32 argc, vbvar *argv, int i, vb_bool d) { return missing(argc, argv, i) ? d : vb_v2bool(argv[i]); }
static vb_i32 along(vb_i32 argc, vbvar *argv, int i, vb_i32 d) { return missing(argc, argv, i) ? d : vb_v2l(argv[i]); }
static vbvar retobj(vbobj o) { return vb_vtmp(vb_v_obj(o)); }

static vb_date unix2date(long long t)
{
    time_t tt = (time_t)t;
    struct tm lt;
    localtime_r(&tt, &lt);
    return vb_dateserial(lt.tm_year + 1900, lt.tm_mon + 1, lt.tm_mday) +
           (lt.tm_hour * 3600 + lt.tm_min * 60 + lt.tm_sec) / 86400.0;
}

static int exists(const char *p, int want_dir)
{
    int a;
    if (vbos_stat(p, 0, 0, &a) != 0) return 0;
    return want_dir < 0 || (want_dir ? (a & 16) != 0 : (a & 16) == 0);
}

/* the parts of a path, as the FileSystemObject sees them */
static const char *last_part(const char *p, int *len)
{
    int n = (int)strlen(p);
    while (n > 1 && is_sep(p[n - 1])) n--;                 /* a final separator does not count */
    int s = n;
    while (s > 0 && !is_sep(p[s - 1]) && p[s - 1] != ':') s--;
    *len = n - s;
    return p + s;
}
static void file_name(const char *p, char *out, int cap)
{
    int l; const char *s = last_part(p, &l);
    if (l == 1 && is_sep(*s)) l = 0;
    snprintf(out, (size_t)cap, "%.*s", l, s);
}
static void base_name(const char *p, char *out, int cap)
{
    file_name(p, out, cap);
    char *dot = strrchr(out, '.');
    if (dot) *dot = 0;
}
static void ext_name(const char *p, char *out, int cap)
{
    char f[1024]; file_name(p, f, sizeof f);
    char *dot = strrchr(f, '.');
    snprintf(out, (size_t)cap, "%s", dot ? dot + 1 : "");
}
static void parent_name(const char *p, char *out, int cap)
{
    int l; const char *s = last_part(p, &l);
    int n = (int)(s - p);
    while (n > 1 && is_sep(p[n - 1]) && !(n == 3 && p[1] == ':')) n--;   /* "C:\" keeps its "\" */
    if (n == 1 && is_sep(p[0]) && l == 0) n = 0;
    snprintf(out, (size_t)cap, "%.*s", n, p);
}
static void join(const char *dir, const char *name, char *out, int cap)
{
    size_t n = strlen(dir);
    char sep = SEP;
    if (strchr(dir, '\\') && !strchr(dir, '/')) sep = '\\';
    else if (strchr(dir, '/')) sep = '/';
    while (*name && is_sep(*name) && n && is_sep(dir[n - 1])) name++;
    if (!n || is_sep(dir[n - 1]) || dir[n - 1] == ':' || is_sep(*name)) snprintf(out, (size_t)cap, "%s%s", dir, name);
    else snprintf(out, (size_t)cap, "%s%c%s", dir, sep, name);
}
static int has_wild(const char *p) { int l; const char *s = last_part(p, &l); return memchr(s, '*', (size_t)l) || memchr(s, '?', (size_t)l); }

/* ---------------------------------------------------------------- File, Folder, Drive objects */
typedef struct { vbobjh h; char *path; } fsitem;
static void fsitem_clear(vbobj o) { free(((fsitem *)o)->path); ((fsitem *)o)->path = 0; }
static vbobj new_item(const vb_class *c, const char *path)
{
    vbobj o = vb_onew(c);
    if (!o) return 0;
    char ab[4096]; vbos_abspath(path, ab, sizeof ab);
    ((fsitem *)o)->path = strdup(ab);
    return o;
}

/* collections: Files / SubFolders of a folder, Drives; the names are read when they are made */
typedef struct { vbobjh h; char *path; int what; int n; char **names; } fscoll;   /* what: 0 files, 1 folders, 2 drives */
static void fscoll_clear(vbobj o)
{
    fscoll *c = (fscoll *)o;
    for (int i = 0; i < c->n; i++) free(c->names[i]);
    free(c->names); free(c->path);
    c->names = 0; c->n = 0; c->path = 0;
}
static int cmp_names(const void *a, const void *b) { return strcasecmp(*(char *const *)a, *(char *const *)b); }
static vbobj new_coll(int what, const char *path)
{
    const vb_class *cls = what == 0 ? &vb_cls_files : what == 1 ? &vb_cls_folders : &vb_cls_drives;
    vbobj o = vb_onew(cls);
    if (!o) return 0;
    fscoll *c = (fscoll *)o;
    c->what = what;
    c->path = strdup(path ? path : "");
    int cap = 0;
    if (what == 2) {
        char buf[8192]; vbos_drives(buf, sizeof buf);
        for (char *p = buf; *p; p += strlen(p) + 1) {
            if (c->n == cap) { cap = cap ? cap * 2 : 16; c->names = (char **)realloc(c->names, sizeof(char *) * (size_t)cap); }
            c->names[c->n++] = strdup(p);
        }
        return o;
    }
    vbos_dir *d = vbos_opendir(path);
    if (!d) return o;
    const char *nm; int isdir;
    while ((nm = vbos_readdir(d, &isdir))) {
        char full[4096]; join(path, nm, full, sizeof full);
        int a = 0;
        if (vbos_stat(full, 0, 0, &a) == 0) isdir = (a & 16) != 0;
        if (isdir != (what == 1)) continue;
        if (c->n == cap) { cap = cap ? cap * 2 : 16; c->names = (char **)realloc(c->names, sizeof(char *) * (size_t)cap); }
        c->names[c->n++] = strdup(nm);
    }
    vbos_closedir(d);
    if (c->n > 1) qsort(c->names, (size_t)c->n, sizeof(char *), cmp_names);
    return o;
}
static vbobj coll_obj(fscoll *c, int i)
{
    if (c->what == 2) return new_item(&vb_cls_drive, c->names[i]);
    char full[4096]; join(c->path, c->names[i], full, sizeof full);
    return new_item(c->what ? &vb_cls_folder : &vb_cls_file, full);
}
static vb_i32 fscoll_each_count(vbobj o) { return ((fscoll *)o)->n; }
static vbvar fscoll_each_item(vbobj o, vb_i32 i)
{
    fscoll *c = (fscoll *)o;
    if (i < 0 || i >= c->n) return vb_v_empty();
    return retobj(coll_obj(c, i));
}
static vbobj fso_create_folder(const char *p);
static vbvar fscoll_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    fscoll *c = (fscoll *)vb_oid(o);
    (void)kind;
    *found = 1;
    if (!strcmp(name, "count")) return vb_v_l(c->n);
    if (!strcmp(name, "item")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        char *key = astr(argv[0]);
        for (int i = 0; i < c->n; i++) {
            const char *nm = c->names[i];
            if (c->what == 2) {                 /* Drives("C"), Drives("C:") */
                if (!strcasecmp(nm, key) || (strlen(key) == 1 && toupper((unsigned char)nm[0]) == toupper((unsigned char)key[0]) && nm[1] == ':'))
                    return retobj(coll_obj(c, i));
            } else if (!strcasecmp(nm, key)) return retobj(coll_obj(c, i));
        }
        vb_raise(c->what == 2 ? 68 : c->what == 1 ? 76 : 53);
        return vb_v_empty();
    }
    if (!strcmp(name, "add") && c->what == 1) {             /* Folders.Add name */
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        char full[4096]; join(c->path, apath(argv[0]), full, sizeof full);
        vbobj f = fso_create_folder(full);
        return f ? retobj(f) : vb_v_empty();
    }
    if (!strcmp(name, "_newenum")) return retobj(o);
    *found = 0;
    return vb_v_empty();
}

/* size of a folder: everything in it */
static double tree_size(const char *path, int depth)
{
    double s = 0;
    if (depth > 64) return 0;
    vbos_dir *d = vbos_opendir(path);
    if (!d) return 0;
    const char *nm; int isdir;
    while ((nm = vbos_readdir(d, &isdir))) {
        char full[4096]; join(path, nm, full, sizeof full);
        long long sz; int a;
        if (vbos_stat(full, &sz, 0, &a) != 0) continue;
        s += (a & 16) ? tree_size(full, depth + 1) : (double)sz;
    }
    vbos_closedir(d);
    return s;
}

static int copy_tree(const char *from, const char *to, int overwrite, int depth);
static int delete_tree(const char *path, int depth);
static vbobj open_text(const char *path, int mode, int create, int format, int truncate);
static const char *drive_of_path(const char *p, char *out, int cap);

/* members that File and Folder have in common */
static vbvar item_common(fsitem *it, int is_folder, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    const char *p = it->path;
    char b[4096];
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    if (!strcmp(name, "path")) return vstr(p);
    if (!strcmp(name, "name")) {
        if (set) {                              /* rename */
            char par[4096], to[4096];
            parent_name(p, par, sizeof par);
            join(par, apath(argv[argc - 1]), to, sizeof to);
            if (exists(to, -1)) { vb_raise(58); return vb_v_empty(); }
            if (vbos_rename(p, to) != 0) { vb_raise(70); return vb_v_empty(); }
            free(it->path); it->path = strdup(to);
            return vb_v_empty();
        }
        file_name(p, b, sizeof b);
        return vstr(*b ? b : p);
    }
    if (!strcmp(name, "shortname")) { file_name(p, b, sizeof b); return vstr(b); }
    if (!strcmp(name, "shortpath")) return vstr(p);
    if (!strcmp(name, "attributes")) {
        if (set) { if (vbos_setattr(p, vb_v2l(argv[argc - 1])) != 0) vb_raise(70); return vb_v_empty(); }
        int a = 0;
        if (vbos_stat(p, 0, 0, &a) != 0) { vb_raise(is_folder ? 76 : 53); return vb_v_empty(); }
        if (!is_folder && !(a & 16)) a |= 32;   /* Archive, as Windows reports most files */
        return vb_v_l(a);
    }
    if (!strcmp(name, "datecreated") || !strcmp(name, "datelastaccessed") || !strcmp(name, "datelastmodified")) {
        long long c, a, m;
        if (vbos_times(p, &c, &a, &m) != 0) { vb_raise(is_folder ? 76 : 53); return vb_v_empty(); }
        return vb_v_date(unix2date(name[4] == 'c' ? c : name[8] == 'a' ? a : m));
    }
    if (!strcmp(name, "parentfolder")) {
        parent_name(p, b, sizeof b);
        if (!*b || !strcmp(b, p)) return retobj(0);
        return retobj(new_item(&vb_cls_folder, b));
    }
    if (!strcmp(name, "drive")) {
        char d[1024]; drive_of_path(p, d, sizeof d);
        return retobj(new_item(&vb_cls_drive, d));
    }
    if (!strcmp(name, "delete")) {
        int r = is_folder ? delete_tree(p, 0) : vbos_remove(p);
        if (r != 0) vb_raise(70);
        return vb_v_empty();
    }
    if (!strcmp(name, "copy") || !strcmp(name, "move")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        char to[4096]; snprintf(to, sizeof to, "%s", apath(argv[0]));
        int over = abool(argc, argv, 1, -1) != 0;
        size_t n = strlen(to);
        if (n && is_sep(to[n - 1])) { char f[1024]; file_name(p, f, sizeof f); strncat(to, f, sizeof to - n - 1); }
        if (name[0] == 'm') {
            if (exists(to, -1)) { vb_raise(58); return vb_v_empty(); }
            if (vbos_rename(p, to) != 0) { vb_raise(70); return vb_v_empty(); }
            free(it->path); it->path = strdup(to);
            return vb_v_empty();
        }
        if (is_folder) { if (copy_tree(p, to, over, 0) != 0 && !vb_errpend) vb_raise(70); }
        else {
            if (!over && exists(to, -1)) { vb_raise(58); return vb_v_empty(); }
            if (vbos_copyfile(p, to) != 0) vb_raise(70);
        }
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

static vbvar file_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    fsitem *it = (fsitem *)vb_oid(o);
    vbvar r = item_common(it, 0, name, kind, argc, argv, found);
    if (*found || vb_errpend) return r;
    *found = 1;
    if (!strcmp(name, "size")) {
        long long sz;
        if (vbos_stat(it->path, &sz, 0, 0) != 0) { vb_raise(53); return vb_v_empty(); }
        return sz <= 2147483647LL ? vb_v_l((vb_i32)sz) : vb_v_d((double)sz);
    }
    if (!strcmp(name, "type")) {
        char e[256]; ext_name(it->path, e, sizeof e);
        char t[300];
        for (char *q = e; *q; q++) *q = (char)toupper((unsigned char)*q);
        snprintf(t, sizeof t, *e ? "%s File" : "File", e);
        return vstr(t);
    }
    if (!strcmp(name, "openastextstream")) {
        vbobj ts = open_text(it->path, along(argc, argv, 0, 1), 0, along(argc, argv, 1, 0), 0);
        return ts ? retobj(ts) : vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

static vbvar folder_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    fsitem *it = (fsitem *)vb_oid(o);
    vbvar r = item_common(it, 1, name, kind, argc, argv, found);
    if (*found || vb_errpend) return r;
    *found = 1;
    if (!strcmp(name, "size")) return vb_v_d(tree_size(it->path, 0));
    if (!strcmp(name, "type")) return vstr("File folder");
    if (!strcmp(name, "files")) return retobj(new_coll(0, it->path));
    if (!strcmp(name, "subfolders")) return retobj(new_coll(1, it->path));
    if (!strcmp(name, "isrootfolder")) {
        char d[1024]; drive_of_path(it->path, d, sizeof d);
        return vb_v_bool(!strcmp(d, it->path) || (strlen(it->path) <= 3 && it->path[1] == ':') ? -1 : 0);
    }
    if (!strcmp(name, "createtextfile")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        char full[4096]; join(it->path, apath(argv[0]), full, sizeof full);
        if (!abool(argc, argv, 1, -1) && exists(full, -1)) { vb_raise(58); return vb_v_empty(); }
        vbobj ts = open_text(full, 2, 1, abool(argc, argv, 2, 0) ? -1 : 0, 1);
        return ts ? retobj(ts) : vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* the drive (root) of a path: "C:" on Windows, the mount point on Linux */
static const char *drive_of_path(const char *p, char *out, int cap)
{
#ifdef _WIN32
    char ab[4096]; vbos_abspath(p, ab, sizeof ab);
    if (ab[0] && ab[1] == ':') snprintf(out, (size_t)cap, "%c:", toupper((unsigned char)ab[0]));
    else if (is_sep(ab[0]) && is_sep(ab[1])) {           /* \\server\share */
        const char *s = ab + 2, *e = strpbrk(s, "\\/");
        if (e) { const char *e2 = strpbrk(e + 1, "\\/"); snprintf(out, (size_t)cap, "%.*s", (int)(e2 ? e2 - ab : (long)strlen(ab)), ab); }
        else snprintf(out, (size_t)cap, "%s", ab);
    } else snprintf(out, (size_t)cap, "%s", ab);
#else
    char ab[4096]; vbos_abspath(p, ab, sizeof ab);
    char buf[8192]; vbos_drives(buf, sizeof buf);
    const char *best = "/";
    for (char *q = buf; *q; q += strlen(q) + 1) {
        size_t l = strlen(q);
        if (!strncmp(ab, q, l) && (ab[l] == '/' || ab[l] == 0) && l > strlen(best)) best = q;
    }
    snprintf(out, (size_t)cap, "%s", best);
#endif
    return out;
}

static vbvar drive_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    fsitem *it = (fsitem *)vb_oid(o);
    const char *p = it->path;
    char root[1100];
#ifdef _WIN32
    snprintf(root, sizeof root, "%s%s", p, p[0] && p[1] == ':' && !p[2] ? "\\" : "");
#else
    snprintf(root, sizeof root, "%s", p);
#endif
    (void)argc; (void)argv;
    *found = 1;
    if (kind == VBM_LET && !strcmp(name, "volumename")) return vb_v_empty();    /* not changed */
    if (!strcmp(name, "path")) return vstr(p);
    if (!strcmp(name, "driveletter")) {
        char l[2] = { 0, 0 };
        if (p[0] && p[1] == ':') l[0] = (char)toupper((unsigned char)p[0]);
        return vstr(l);
    }
    if (!strcmp(name, "rootfolder")) return retobj(new_item(&vb_cls_folder, root));
    if (!strcmp(name, "totalsize") || !strcmp(name, "freespace") || !strcmp(name, "availablespace")) {
        double t, f, a;
        if (vbos_disk(root, &t, &f, &a) != 0) { vb_raise(71); return vb_v_empty(); }   /* Disk not ready */
        return vb_v_d(name[0] == 't' ? t : name[0] == 'f' ? f : a);
    }
    char vol[512], fs[128]; unsigned ser; int type;
    int ok = vbos_drive_info(root, vol, sizeof vol, fs, sizeof fs, &ser, &type) == 0;
    if (!strcmp(name, "isready")) return vb_v_bool(ok && vbos_disk(root, 0, 0, 0) == 0 ? -1 : 0);
    if (!strcmp(name, "drivetype")) return vb_v_l(type);
    if (!strcmp(name, "filesystem")) return vstr(fs);
    if (!strcmp(name, "volumename")) return vstr(vol);
    if (!strcmp(name, "serialnumber")) return vb_v_l((vb_i32)ser);
    if (!strcmp(name, "sharename")) return vstr(type == 3 ? p : "");
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- TextStream
 * Reading: the whole file is read when it is opened (ANSI in the code page of the program, or
 * Unicode UTF-16 with or without a BOM); writing goes to the file at once. */
typedef struct {
    vbobjh h;
    FILE *f;
    int mode;                                   /* 1 ForReading, 2 ForWriting, 8 ForAppending */
    int unicode;
    vb_wchar *buf; vb_i32 n, pos;               /* reading */
    vb_i32 line, col;
} tstream;
static void ts_clear(vbobj o)
{
    tstream *t = (tstream *)o;
    if (t->f) fclose(t->f);
    free(t->buf);
    t->f = 0; t->buf = 0;
}
static vbobj open_text(const char *path, int mode, int create, int format, int truncate)
{
    if (mode != 1 && mode != 2 && mode != 8) { vb_raise(5); return 0; }
    int ex = exists(path, 0);
    if (exists(path, 1)) { vb_raise(70); return 0; }      /* a folder: Permission denied */
    if (!ex && !create && mode == 1) { vb_raise(53); return 0; }
    if (!ex && !create && mode != 1 && !truncate) { vb_raise(53); return 0; }
    int unicode = format == -1;
    if (mode == 1) {
        FILE *f = vbos_fopen(path, "rb");
        if (!f) { if (ex) { vb_raise(70); return 0; } f = 0; }
        vbobj o = vb_onew(&vb_cls_textstream);
        tstream *t = (tstream *)o;
        t->mode = 1; t->line = 1; t->col = 1;
        long sz = 0; unsigned char *raw = 0;
        if (f) {
            fseek(f, 0, SEEK_END); sz = ftell(f); fseek(f, 0, SEEK_SET);
            raw = (unsigned char *)malloc((size_t)sz + 2);
            sz = (long)fread(raw, 1, (size_t)sz, f);
            fclose(f);
        }
        long start = 0;
        if (sz >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) { unicode = 1; start = 2; }   /* a BOM: Unicode */
        else if (format == -2) unicode = 0;
        if (unicode) {
            t->n = (vb_i32)((sz - start) / 2);
            t->buf = (vb_wchar *)malloc(sizeof(vb_wchar) * ((size_t)t->n + 1));
            for (vb_i32 i = 0; i < t->n; i++) t->buf[i] = (vb_wchar)(raw[start + 2 * i] | (raw[start + 2 * i + 1] << 8));
        } else {
            t->n = (vb_i32)sz;
            t->buf = (vb_wchar *)malloc(sizeof(vb_wchar) * ((size_t)t->n + 1));
            for (vb_i32 i = 0; i < t->n; i++) t->buf[i] = vb_ansi2w(raw[i]);
        }
        t->unicode = unicode;
        free(raw);
        return o;
    }
    FILE *f = vbos_fopen(path, mode == 8 && !truncate ? "ab" : "wb");
    if (!f) { vb_raise(70); return 0; }
    vbobj o = vb_onew(&vb_cls_textstream);
    tstream *t = (tstream *)o;
    t->f = f; t->mode = mode; t->unicode = unicode; t->line = 1; t->col = 1;
    if (unicode && ftell(f) == 0) { fputc(0xFF, f); fputc(0xFE, f); }
    return o;
}
static void ts_write(tstream *t, vbs s)
{
    for (vb_i32 i = 0; i < s.len; i++) {
        vb_wchar c = s.p[i];
        if (t->unicode) { fputc(c & 255, t->f); fputc(c >> 8, t->f); }
        else { int b = vb_w2ansi(c); fputc(b < 0 ? '?' : b, t->f); }
        if (c == '\n') { t->line++; t->col = 1; } else t->col++;
    }
}
static void ts_advance(tstream *t, vb_i32 to)
{
    for (; t->pos < to; t->pos++) {
        if (t->buf[t->pos] == '\n') { t->line++; t->col = 1; } else t->col++;
    }
}
static vbvar ts_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    tstream *t = (tstream *)vb_oid(o);
    (void)kind;
    *found = 1;
    int reading = t->mode == 1;
    if (!strcmp(name, "close")) { if (t->f) fclose(t->f); t->f = 0; free(t->buf); t->buf = 0; t->n = t->pos = 0; t->mode = 0; return vb_v_empty(); }
    if (!strcmp(name, "line")) return vb_v_l(t->line);
    if (!strcmp(name, "column")) return vb_v_l(t->col);
    if (!strcmp(name, "atendofstream")) {
        if (!reading) { vb_raise(54); return vb_v_empty(); }
        return vb_v_bool(t->pos >= t->n ? -1 : 0);
    }
    if (!strcmp(name, "atendofline")) {
        if (!reading) { vb_raise(54); return vb_v_empty(); }
        return vb_v_bool(t->pos >= t->n || t->buf[t->pos] == '\r' || t->buf[t->pos] == '\n' ? -1 : 0);
    }
    if (!strcmp(name, "read") || !strcmp(name, "readline") || !strcmp(name, "readall") || !strcmp(name, "skip") || !strcmp(name, "skipline")) {
        if (!reading) { vb_raise(54); return vb_v_empty(); }            /* Bad file mode */
        if (t->pos >= t->n) { vb_raise(62); return vb_v_empty(); }      /* Input past end of file */
        vb_i32 a = t->pos, b;
        int line = name[0] == 'r' ? !strcmp(name, "readline") : !strcmp(name, "skipline");
        if (!strcmp(name, "readall")) b = t->n;
        else if (line) {
            b = a;
            while (b < t->n && t->buf[b] != '\n' && t->buf[b] != '\r') b++;
            vb_i32 e = b;
            if (e < t->n && t->buf[e] == '\r') e++;
            if (e < t->n && t->buf[e] == '\n') e++;
            ts_advance(t, e);
            if (name[0] == 's') return vb_v_empty();
            vbs s = vb_stmp(b - a); memcpy(s.p, t->buf + a, sizeof(vb_wchar) * (size_t)(b - a));
            return vb_v_s(s);
        } else {
            vb_i32 k = vb_v2l(missing(argc, argv, 0) ? vb_v_l(0) : argv[0]);
            if (vb_errpend) return vb_v_empty();
            if (k < 0) { vb_raise(5); return vb_v_empty(); }
            b = a + k > t->n ? t->n : a + k;
        }
        ts_advance(t, b);
        if (name[0] == 's') return vb_v_empty();
        vbs s = vb_stmp(b - a); memcpy(s.p, t->buf + a, sizeof(vb_wchar) * (size_t)(b - a));
        return vb_v_s(s);
    }
    if (!strcmp(name, "write") || !strcmp(name, "writeline") || !strcmp(name, "writeblanklines")) {
        if (reading || !t->f) { vb_raise(54); return vb_v_empty(); }
        static const vb_wchar crlf[] = { '\r', '\n' };
        vbs nl = vb_slit(crlf, 2);
        if (!strcmp(name, "writeblanklines")) { vb_i32 k = along(argc, argv, 0, 0); while (k-- > 0) ts_write(t, nl); return vb_v_empty(); }
        if (!missing(argc, argv, 0)) { vbs s = vb_v2s(argv[0]); if (vb_errpend) return vb_v_empty(); ts_write(t, s); }
        else if (name[5] != 'l') { vb_raise(450); return vb_v_empty(); }
        if (name[5] == 'l') ts_write(t, nl);
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- FileSystemObject */
static vbobj fso_create_folder(const char *p)
{
    if (exists(p, -1)) { vb_raise(58); return 0; }
    char par[4096]; parent_name(p, par, sizeof par);
    if (*par && !exists(par, 1)) { vb_raise(76); return 0; }
    if (vbos_mkdir(p) != 0) { vb_raise(70); return 0; }
    return new_item(&vb_cls_folder, p);
}

/* a root ("/", "C:\", "\\server\share\"): DeleteFolder / Folder.Delete of it is refused (error 70) and nothing
 * is touched - robot 2026-11-10: Dir2HTML answered "Yes" to "delete / ?" and called DeleteFolder "/" */
static int is_root(const char *path)
{
    char a[4096];
    vbos_abspath(path, a, sizeof a);
    int n = (int)strlen(a);
    while (n > 0 && is_sep(a[n - 1])) n--;
    if (n == 0) return 1;                                         /* "/" */
    if (n == 2 && a[1] == ':') return 1;                          /* "C:" */
    if (n > 2 && is_sep(a[0]) && is_sep(a[1])) {                  /* \\server\share */
        int seps = 0;
        for (int i = 2; i < n; i++) if (is_sep(a[i])) seps++;
        return seps <= 1;
    }
    return 0;
}

static int delete_tree(const char *path, int depth)
{
    if (depth > 64) return -1;
    if (depth == 0 && is_root(path)) return -1;
    if (depth == 0) {                                             /* DeleteFolder of a link: the link only */
        char p[4096]; snprintf(p, sizeof p, "%s", path);
        size_t n = strlen(p);
        while (n > 1 && is_sep(p[n - 1])) p[--n] = 0;            /* ("link/" would be the folder it leads to) */
        if (!vbos_realdir(p)) return vbos_remove(p) == 0 || vbos_rmdir(p) == 0 ? 0 : -1;
    }
    vbos_dir *d = vbos_opendir(path);
    if (d) {
        const char *nm; int isdir;
        char **names = 0; int n = 0, cap = 0;
        while ((nm = vbos_readdir(d, &isdir))) {
            if (n == cap) { cap = cap ? cap * 2 : 16; names = (char **)realloc(names, sizeof(char *) * (size_t)cap); }
            names[n++] = strdup(nm);
        }
        vbos_closedir(d);
        for (int i = 0; i < n; i++) {
            char full[4096]; join(path, names[i], full, sizeof full);
            if (vbos_realdir(full)) delete_tree(full, depth + 1);   /* a link to a folder is not followed */
            else if (vbos_remove(full) != 0) vbos_rmdir(full);       /* (a junction on Windows goes with rmdir) */
            free(names[i]);
        }
        free(names);
    }
    return vbos_rmdir(path);
}

static int copy_tree(const char *from, const char *to, int overwrite, int depth)
{
    if (depth > 64) return -1;
    if (!exists(to, 1) && vbos_mkdir(to) != 0) return -1;
    vbos_dir *d = vbos_opendir(from);
    if (!d) return -1;
    const char *nm; int isdir, r = 0;
    while ((nm = vbos_readdir(d, &isdir))) {
        char a[4096], b[4096];
        join(from, nm, a, sizeof a); join(to, nm, b, sizeof b);
        if (exists(a, 1)) { if (copy_tree(a, b, overwrite, depth + 1) != 0) r = -1; }
        else {
            if (exists(b, -1) && !overwrite) { vb_raise(58); r = -1; break; }
            if (vbos_copyfile(a, b) != 0) r = -1;
        }
    }
    vbos_closedir(d);
    return r;
}

/* CopyFile / MoveFile / DeleteFile / CopyFolder / MoveFolder / DeleteFolder: the last part of the
 * source may have * and ? */
static void fso_bulk(const char *op, const char *src, const char *dst, int overwrite)
{
    int folder = strstr(op, "folder") != 0;
    char dir[4096], pat[1024];
    if (has_wild(src)) { parent_name(src, dir, sizeof dir); file_name(src, pat, sizeof pat); if (!*dir) snprintf(dir, sizeof dir, "."); }
    else { dir[0] = 0; pat[0] = 0; }
    char **list = 0; int n = 0, cap = 0;
    if (*pat) {
        vbos_dir *d = vbos_opendir(dir);
        if (!d) { vb_raise(76); return; }
        const char *nm; int isdir;
        while ((nm = vbos_readdir(d, &isdir))) {
            char full[4096]; join(dir, nm, full, sizeof full);
            if (!vbos_wildmatch(pat, nm) || !exists(full, folder)) continue;
            if (n == cap) { cap = cap ? cap * 2 : 16; list = (char **)realloc(list, sizeof(char *) * (size_t)cap); }
            list[n++] = strdup(full);
        }
        vbos_closedir(d);
        if (!n) { vb_raise(folder ? 76 : 53); free(list); return; }
    } else {
        if (!exists(src, folder)) { vb_raise(folder ? 76 : 53); return; }
        list = (char **)malloc(sizeof(char *)); list[n++] = strdup(src);
    }
    int many = *pat != 0;
    for (int i = 0; i < n && !vb_errpend; i++) {
        const char *s = list[i];
        char to[4096];
        if (dst) {
            size_t l = strlen(dst);
            int into = many || (l && is_sep(dst[l - 1])) || (folder && exists(dst, 1) && op[0] != 'm') || (!folder && exists(dst, 1));
            if (into) { char f[1024]; file_name(s, f, sizeof f); join(dst, f, to, sizeof to); }
            else snprintf(to, sizeof to, "%s", dst);
        }
        if (!strncmp(op, "delete", 6)) {
            int r = folder ? delete_tree(s, 0) : vbos_remove(s);
            if (r != 0) vb_raise(70);
        } else if (!strncmp(op, "move", 4)) {
            if (exists(to, -1)) vb_raise(58);
            else if (vbos_rename(s, to) != 0) vb_raise(70);
        } else if (folder) {
            if (copy_tree(s, to, overwrite, 0) != 0 && !vb_errpend) vb_raise(70);
        } else {
            if (!overwrite && exists(to, -1)) vb_raise(58);
            else if (vbos_copyfile(s, to) != 0) vb_raise(70);
        }
    }
    for (int i = 0; i < n; i++) free(list[i]);
    free(list);
}

static vbvar fso_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)o; (void)kind;
    char b[4096];
    *found = 1;
    if (!strcmp(name, "drives")) return retobj(new_coll(2, 0));
    if (argc < 1 && strcmp(name, "gettempname") != 0) {
        static const char *const known[] = { "buildpath", "copyfile", "copyfolder", "createfolder", "createtextfile",
            "deletefile", "deletefolder", "driveexists", "fileexists", "folderexists", "getabsolutepathname",
            "getbasename", "getdrive", "getdrivename", "getextensionname", "getfile", "getfilename", "getfolder",
            "getparentfoldername", "getspecialfolder", "movefile", "movefolder", "opentextfile", "getfileversion",
            "getstandardstream", 0 };
        for (int i = 0; known[i]; i++) if (!strcmp(name, known[i])) { vb_raise(450); return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
    if (!strcmp(name, "buildpath")) {
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        char d[4096]; snprintf(d, sizeof d, "%s", astr(argv[0]));
        join(d, astr(argv[1]), b, sizeof b);
        return vstr(b);
    }
    if (!strcmp(name, "getfilename")) { file_name(astr(argv[0]), b, sizeof b); return vstr(b); }
    if (!strcmp(name, "getbasename")) { base_name(astr(argv[0]), b, sizeof b); return vstr(b); }
    if (!strcmp(name, "getextensionname")) { ext_name(astr(argv[0]), b, sizeof b); return vstr(b); }
    if (!strcmp(name, "getparentfoldername")) { parent_name(astr(argv[0]), b, sizeof b); return vstr(b); }
    if (!strcmp(name, "getabsolutepathname")) { vbos_abspath(apath(argv[0]), b, sizeof b); return vstr(b); }
    if (!strcmp(name, "getdrivename")) {
        const char *p = astr(argv[0]);
        if (p[0] && p[1] == ':') { snprintf(b, sizeof b, "%c:", p[0]); return vstr(b); }
        return vstr("");
    }
    if (!strcmp(name, "fileexists")) return vb_v_bool(exists(apath(argv[0]), 0) ? -1 : 0);
    if (!strcmp(name, "folderexists")) return vb_v_bool(exists(apath(argv[0]), 1) ? -1 : 0);
    if (!strcmp(name, "driveexists")) {
        const char *p = apath(argv[0]);
        char buf[8192]; vbos_drives(buf, sizeof buf);
        for (char *q = buf; *q; q += strlen(q) + 1)
            if (!strcasecmp(q, p) || (strlen(p) <= 3 && p[0] && toupper((unsigned char)p[0]) == toupper((unsigned char)q[0]) && q[1] == ':'))
                return vb_v_bool(-1);
        return vb_v_bool(exists(p, 1) && !strcmp(drive_of_path(p, b, sizeof b), p) ? -1 : 0);
    }
    if (!strcmp(name, "getfile")) {
        const char *p = apath(argv[0]);
        if (!exists(p, 0)) { vb_raise(53); return vb_v_empty(); }
        return retobj(new_item(&vb_cls_file, p));
    }
    if (!strcmp(name, "getfolder")) {
        const char *p = apath(argv[0]);
        if (!exists(p, 1)) { vb_raise(76); return vb_v_empty(); }
        return retobj(new_item(&vb_cls_folder, p));
    }
    if (!strcmp(name, "getdrive")) {
        const char *p = apath(argv[0]);
        char buf[8192]; vbos_drives(buf, sizeof buf);
        for (char *q = buf; *q; q += strlen(q) + 1)
            if (!strcasecmp(q, p) || (strlen(p) <= 3 && p[0] && toupper((unsigned char)p[0]) == toupper((unsigned char)q[0]) && q[1] == ':'))
                return retobj(new_item(&vb_cls_drive, q));
        if (exists(p, 1)) return retobj(new_item(&vb_cls_drive, drive_of_path(p, b, sizeof b)));
        vb_raise(68);                           /* Device unavailable */
        return vb_v_empty();
    }
    if (!strcmp(name, "getspecialfolder")) {
        vb_i32 w = vb_v2l(argv[0]);
        if (vb_errpend) return vb_v_empty();
        if (w < 0 || w > 2) { vb_raise(5); return vb_v_empty(); }
        vbos_special_folder(w, b, sizeof b);
        return retobj(new_item(&vb_cls_folder, b));
    }
    if (!strcmp(name, "createfolder")) { vbobj f = fso_create_folder(apath(argv[0])); return f ? retobj(f) : vb_v_empty(); }
    if (!strcmp(name, "createtextfile")) {
        const char *p = apath(argv[0]);
        if (!abool(argc, argv, 1, -1) && exists(p, -1)) { vb_raise(58); return vb_v_empty(); }
        vbobj t = open_text(p, 2, 1, abool(argc, argv, 2, 0) ? -1 : 0, 1);
        return t ? retobj(t) : vb_v_empty();
    }
    if (!strcmp(name, "opentextfile")) {
        vbobj t = open_text(apath(argv[0]), along(argc, argv, 1, 1), abool(argc, argv, 2, 0) != 0, along(argc, argv, 3, 0), 0);
        return t ? retobj(t) : vb_v_empty();
    }
    if (!strcmp(name, "copyfile") || !strcmp(name, "copyfolder") || !strcmp(name, "movefile") || !strcmp(name, "movefolder")) {
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        char s[4096]; snprintf(s, sizeof s, "%s", apath(argv[0]));
        fso_bulk(name, s, apath(argv[1]), abool(argc, argv, 2, -1) != 0);
        return vb_v_empty();
    }
    if (!strcmp(name, "deletefile") || !strcmp(name, "deletefolder")) { fso_bulk(name, apath(argv[0]), 0, 1); return vb_v_empty(); }
    if (!strcmp(name, "gettempname")) {
        static int seeded;
        if (!seeded) { srand((unsigned)time(0) ^ (unsigned)(size_t)&seeded); seeded = 1; }
        snprintf(b, sizeof b, "rad%05X.tmp", rand() & 0xFFFFF);
        return vstr(b);
    }
    if (!strcmp(name, "getfileversion")) return vstr("");
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Dictionary
 * Keys in the order they were added, with a hash index; Item of a key that is not there adds it
 * (as scrrun.dll does); For Each goes over the keys. */
typedef struct {
    vbobjh h;
    vb_i32 n, cap, cmp;                         /* cmp: CompareMode 0 binary, 1 text */
    vbvar *keys, *items;
    vb_i32 *slots; vb_i32 nslots;               /* index + 1, 0 = empty */
} dict;
static void dict_clear(vbobj o)
{
    dict *d = (dict *)o;
    for (vb_i32 i = 0; i < d->n; i++) { vb_vfree(&d->keys[i]); vb_vfree(&d->items[i]); }
    free(d->keys); free(d->items); free(d->slots);
    d->keys = d->items = 0; d->slots = 0; d->n = d->cap = d->nslots = 0;
}
static int numeric(vbvar v)
{
    int t = v.vt;
    return t == VT_I2 || t == VT_I4 || t == VT_R4 || t == VT_R8 || t == VT_CY || t == VT_DATE || t == VT_BOOL || t == VT_UI1;
}
static unsigned key_hash(dict *d, vbvar k)
{
    unsigned h = 2166136261u;
    if (k.vt == VT_BSTR) {
        for (vb_i32 i = 0; i < k.u.s.len; i++) {
            vb_wchar c = k.u.s.p[i];
            if (d->cmp && c < 128) c = (vb_wchar)tolower(c);
            else if (d->cmp) c = 0;             /* text: other letters compare by locale; one bucket */
            h = (h ^ c) * 16777619u;
        }
        return h;
    }
    if (numeric(k)) { double x = vb_v2d(k); if (x == 0) x = 0; unsigned char *b = (unsigned char *)&x; for (int i = 0; i < 8; i++) h = (h ^ b[i]) * 16777619u; return h ^ 1; }
    if (k.vt == VT_DISPATCH) return (unsigned)(((size_t)vb_oid(k.u.o)) >> 4);
    return (unsigned)k.vt;
}
static int key_eq(dict *d, vbvar a, vbvar b)
{
    if (a.vt == VT_BSTR && b.vt == VT_BSTR) return vb_scmp(a.u.s, b.u.s, d->cmp) == 0;
    if (numeric(a) && numeric(b)) return vb_v2d(a) == vb_v2d(b);
    if (a.vt == VT_DISPATCH && b.vt == VT_DISPATCH) return vb_oid(a.u.o) == vb_oid(b.u.o);
    if (a.vt != b.vt) return 0;
    return a.vt == VT_EMPTY || a.vt == VT_NULL;
}
static void dict_rehash(dict *d)
{
    vb_i32 ns = 16;
    while (ns < d->n * 2 + 2) ns *= 2;
    free(d->slots);
    d->slots = (vb_i32 *)calloc((size_t)ns, sizeof(vb_i32));
    d->nslots = ns;
    for (vb_i32 i = 0; i < d->n; i++) {
        unsigned h = key_hash(d, d->keys[i]) & (unsigned)(ns - 1);
        while (d->slots[h]) h = (h + 1) & (unsigned)(ns - 1);
        d->slots[h] = i + 1;
    }
}
static vb_i32 dict_find(dict *d, vbvar k)
{
    if (!d->nslots) return -1;
    unsigned h = key_hash(d, k) & (unsigned)(d->nslots - 1);
    while (d->slots[h]) {
        vb_i32 i = d->slots[h] - 1;
        if (key_eq(d, d->keys[i], k)) return i;
        h = (h + 1) & (unsigned)(d->nslots - 1);
    }
    return -1;
}
static vbvar plain_key(vbvar k)
{
    if (k.vt == VT_ERROR && vb_ismissing(k)) return vb_v_empty();
    return k;
}
static vb_i32 dict_add(dict *d, vbvar k, vbvar item)
{
    if (k.vt & VT_ARRAY) { vb_raise(13); return -1; }
    if (d->n == d->cap) {
        vb_i32 nc = d->cap ? d->cap * 2 : 8;
        d->keys = (vbvar *)realloc(d->keys, sizeof(vbvar) * (size_t)nc);
        d->items = (vbvar *)realloc(d->items, sizeof(vbvar) * (size_t)nc);
        d->cap = nc;
    }
    memset(&d->keys[d->n], 0, sizeof(vbvar)); memset(&d->items[d->n], 0, sizeof(vbvar));
    vb_vset(&d->keys[d->n], k);
    vb_vset(&d->items[d->n], item);
    d->n++;
    if (d->n * 2 + 2 > d->nslots) dict_rehash(d);
    else {
        unsigned h = key_hash(d, k) & (unsigned)(d->nslots - 1);
        while (d->slots[h]) h = (h + 1) & (unsigned)(d->nslots - 1);
        d->slots[h] = d->n;
    }
    return d->n - 1;
}
static vbvar dict_array(dict *d, int items)
{
    vbarr *a = vb_atmp();
    if (d->n) vb_adim(a, sizeof(vbvar), VBA_VAR, 0, 1, 0, d->n - 1);
    else vb_aempty(a, sizeof(vbvar), VBA_VAR);              /* UBound -1 */
    for (vb_i32 i = 0; i < d->n; i++) vb_vset((vbvar *)vb_aelem(a, 1, i), items ? d->items[i] : d->keys[i]);
    vbvar v = vb_v_empty();
    v.vt = VT_ARRAY | VT_VARIANT; v.u.a = a;
    return v;
}
static vb_i32 dict_each_count(vbobj o) { return ((dict *)o)->n; }
static vbvar dict_each_item(vbobj o, vb_i32 i)
{
    dict *d = (dict *)o;
    return i >= 0 && i < d->n ? vb_vtmp(d->keys[i]) : vb_v_empty();
}
static vbvar dict_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dict *d = (dict *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    if (!strcmp(name, "count")) return vb_v_l(d->n);
    if (!strcmp(name, "add")) {
        if (argc != 2) { vb_raise(450); return vb_v_empty(); }
        vbvar k = plain_key(argv[0]);
        if (dict_find(d, k) >= 0) { vb_raise(457); return vb_v_empty(); }   /* key already associated */
        dict_add(d, k, argv[1]);
        return vb_v_empty();
    }
    if (!strcmp(name, "item")) {
        if (argc < (set ? 2 : 1)) { vb_raise(450); return vb_v_empty(); }
        vbvar k = plain_key(argv[0]);
        vb_i32 i = dict_find(d, k);
        if (set) {
            if (i < 0) dict_add(d, k, argv[argc - 1]);
            else vb_vset(&d->items[i], argv[argc - 1]);
            return vb_v_empty();
        }
        if (i < 0) i = dict_add(d, k, vb_v_empty());       /* reading a new key adds it */
        return i < 0 ? vb_v_empty() : vb_vtmp(d->items[i]);
    }
    if (!strcmp(name, "key")) {                 /* d.Key(old) = new */
        if (!set || argc != 2) { vb_raise(set ? 450 : 451); return vb_v_empty(); }
        vbvar k = plain_key(argv[0]), nk = plain_key(argv[1]);
        vb_i32 i = dict_find(d, k);
        if (i < 0) { vb_raise_full(32811, vb_sfromutf8("Scripting.Dictionary"), vb_stmp(0)); return vb_v_empty(); }   /* (no text of its own) */
        if (dict_find(d, nk) >= 0) { vb_raise(457); return vb_v_empty(); }
        vb_vset(&d->keys[i], nk);
        dict_rehash(d);
        return vb_v_empty();
    }
    if (!strcmp(name, "exists")) {
        if (argc != 1) { vb_raise(450); return vb_v_empty(); }
        return vb_v_bool(dict_find(d, plain_key(argv[0])) >= 0 ? -1 : 0);
    }
    if (!strcmp(name, "keys") || !strcmp(name, "items")) return dict_array(d, name[0] == 'i');
    if (!strcmp(name, "remove")) {
        if (argc != 1) { vb_raise(450); return vb_v_empty(); }
        vb_i32 i = dict_find(d, plain_key(argv[0]));
        if (i < 0) { vb_raise_full(32811, vb_sfromutf8("Scripting.Dictionary"), vb_stmp(0)); return vb_v_empty(); }
        vbvar ok = d->keys[i], oi = d->items[i];
        memmove(d->keys + i, d->keys + i + 1, sizeof(vbvar) * (size_t)(d->n - i - 1));
        memmove(d->items + i, d->items + i + 1, sizeof(vbvar) * (size_t)(d->n - i - 1));
        d->n--;
        dict_rehash(d);
        vb_vfree(&ok); vb_vfree(&oi);           /* last: may run Class_Terminate */
        return vb_v_empty();
    }
    if (!strcmp(name, "removeall")) { dict_clear(o); return vb_v_empty(); }
    if (!strcmp(name, "comparemode")) {
        if (!set) return vb_v_l(d->cmp);
        if (d->n) { vb_raise(5); return vb_v_empty(); }
        vb_i32 m = vb_v2l(argv[argc - 1]);
        if (vb_errpend) return vb_v_empty();
        if (m < 0) { vb_raise(5); return vb_v_empty(); }
        d->cmp = m >= 1;                        /* vbTextCompare, vbDatabaseCompare: without case */
        return vb_v_empty();
    }
    if (!strcmp(name, "hashval")) return vb_v_l((vb_i32)(key_hash(d, plain_key(argc ? argv[0] : vb_v_empty())) % 1201));
    if (!strcmp(name, "_newenum")) return retobj(o);
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- the classes */
#define FB(f) 0, 0, 0, f
const vb_class vb_cls_fso = { "FileSystemObject", sizeof(vbobjh), 0, 0, 0, 0, 0, 0, FB(fso_member) };
const vb_class vb_cls_drive = { "Drive", sizeof(fsitem), 0, 0, fsitem_clear, 0, 0, "path", FB(drive_member) };
const vb_class vb_cls_folder = { "Folder", sizeof(fsitem), 0, 0, fsitem_clear, 0, 0, "path", FB(folder_member) };
const vb_class vb_cls_file = { "File", sizeof(fsitem), 0, 0, fsitem_clear, 0, 0, "path", FB(file_member) };
const vb_class vb_cls_files = { "Files", sizeof(fscoll), 0, 0, fscoll_clear, 0, 0, "item", fscoll_each_count, fscoll_each_item, 0, fscoll_member };
const vb_class vb_cls_folders = { "Folders", sizeof(fscoll), 0, 0, fscoll_clear, 0, 0, "item", fscoll_each_count, fscoll_each_item, 0, fscoll_member };
const vb_class vb_cls_drives = { "Drives", sizeof(fscoll), 0, 0, fscoll_clear, 0, 0, "item", fscoll_each_count, fscoll_each_item, 0, fscoll_member };
const vb_class vb_cls_textstream = { "TextStream", sizeof(tstream), 0, 0, ts_clear, 0, 0, 0, FB(ts_member) };
const vb_class vb_cls_dictionary = { "Dictionary", sizeof(dict), 0, 0, dict_clear, 0, 0, "item", dict_each_count, dict_each_item, 0, dict_member };

/* ---------------------------------------------------------------- CreateObject
 * The classes a program can create by name (ProgID). The ActiveX projects of phase 4.3 add theirs. */
static const struct { const char *progid; const vb_class *cls; } progids[] = {
    { "scripting.filesystemobject", &vb_cls_fso },
    { "scripting.dictionary", &vb_cls_dictionary },
};
vbobj vb_createobject(vbs progid, vbvar server)
{
    (void)server;
    char *p = vb_stoutf8(progid);
    char low[256]; int k = 0;
    for (; p[k] && k < 255; k++) low[k] = (char)tolower((unsigned char)p[k]);
    low[k] = 0;
    char *ver = strrchr(low, '.');              /* "Scripting.Dictionary.1": the version does not matter */
    if (ver && ver != strchr(low, '.') && isdigit((unsigned char)ver[1])) *ver = 0;
    for (size_t i = 0; i < sizeof progids / sizeof *progids; i++)
        if (!strcmp(low, progids[i].progid)) return vb_onew(progids[i].cls);
    if (vb_dbhooks.createobject) {              /* DAO.DBEngine.36 ... (libvbdb) */
        vbobj d = vb_dbhooks.createobject(low);
        if (d || vb_errpend) return d;
    }
    int found;
    vbobj o = vb_component_create(low, &found);  /* a component compiled into the program */
    if (found || vb_errpend) return o;
    vb_raise(429);                              /* ActiveX component can't create object */
    return 0;
}
