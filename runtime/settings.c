/* settings.c - GetSetting, SaveSetting, DeleteSetting.
 * VB6 keeps them in the registry under HKEY_CURRENT_USER\Software\VB and VBA Program Settings;
 * here each application has a text file of [Section] Key=Value lines:
 *   Linux:   $XDG_CONFIG_HOME/VB and VBA Program Settings/<AppName>.ini (~/.config/...)
 *   Windows: %APPDATA%\VB and VBA Program Settings\<AppName>.ini
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include "vbos.h"
#include <errno.h>
#include <strings.h>
#include <sys/stat.h>
#ifdef _WIN32
#define SEP "\\"
#else
#define SEP "/"
#endif

typedef struct { char **line; int n, cap; } lines_t;

static void add_line(lines_t *L, const char *s, int at)
{
    if (L->n == L->cap) { L->cap = L->cap ? L->cap * 2 : 32; L->line = (char **)realloc(L->line, L->cap * sizeof(char *)); }
    if (at < 0 || at > L->n) at = L->n;
    memmove(L->line + at + 1, L->line + at, (L->n - at) * sizeof(char *));
    L->line[at] = strdup(s);
    L->n++;
}

static void del_line(lines_t *L, int at)
{
    free(L->line[at]);
    memmove(L->line + at, L->line + at + 1, (L->n - at - 1) * sizeof(char *));
    L->n--;
}

static void free_lines(lines_t *L) { for (int i = 0; i < L->n; i++) free(L->line[i]); free(L->line); }

static char *file_of(const char *app, int make_dir)
{
    static char path[4096];
    char dir[3800];
#ifdef _WIN32
    const char *base = getenv("APPDATA");
    snprintf(dir, sizeof dir, "%s\\VB and VBA Program Settings", base ? base : ".");
#else
    const char *x = getenv("XDG_CONFIG_HOME"), *h = getenv("HOME");
    if (x && *x) snprintf(dir, sizeof dir, "%s/VB and VBA Program Settings", x);
    else snprintf(dir, sizeof dir, "%s/.config/VB and VBA Program Settings", h ? h : ".");
#endif
    if (make_dir)                                  /* mkdir -p */
        for (char *p = dir + 1; ; p++)
            if (*p == '/' || *p == '\\' || *p == 0) {
                char c = *p; *p = 0; vbos_mkdir(dir); *p = c;
                if (!c) break;
            }
    char name[256]; int k = 0;
    for (const char *p = app; *p && k < 250; p++) name[k++] = (*p == '/' || *p == '\\') ? '_' : *p;
    name[k] = 0;
    snprintf(path, sizeof path, "%s" SEP "%s.ini", dir, name);
    return path;
}

static void load(const char *path, lines_t *L)
{
    memset(L, 0, sizeof *L);
    FILE *f = vbos_fopen(path, "rb");
    if (!f) return;
    char buf[8192];
    while (fgets(buf, sizeof buf, f)) {
        size_t n = strlen(buf);
        while (n && (buf[n - 1] == '\n' || buf[n - 1] == '\r')) buf[--n] = 0;
        add_line(L, buf, -1);
    }
    fclose(f);
}

static void save(const char *path, lines_t *L)
{
    FILE *f = vbos_fopen(path, "wb");
    if (!f) { vb_raise(75); return; }
    for (int i = 0; i < L->n; i++) { fputs(L->line[i], f); fputc('\n', f); }
    fclose(f);
}

static int is_section(const char *s, const char *sec)
{
    if (s[0] != '[') return 0;
    const char *e = strrchr(s, ']');
    if (!e) return 0;
    size_t n = (size_t)(e - s - 1);
    return sec == 0 || (strlen(sec) == n && strncasecmp(s + 1, sec, n) == 0);
}

/* a Key=Value line of the key (blanks around the key allowed, case does not matter) */
static int key_is(const char *line, const char *key)
{
    while (*line == ' ' || *line == '\t') line++;
    size_t kn = strlen(key);
    if (strncasecmp(line, key, kn) != 0) return 0;
    line += kn;
    while (*line == ' ' || *line == '\t') line++;
    return *line == '=';
}

/* the line of [sec] (-1), and of key in it (-1); *end: where the section ends */
static int find(lines_t *L, const char *sec, const char *key, int *sec_line, int *end)
{
    *sec_line = -1; *end = L->n;
    for (int i = 0; i < L->n; i++) {
        if (*sec_line < 0) { if (is_section(L->line[i], sec)) *sec_line = i; continue; }
        if (is_section(L->line[i], 0)) { *end = i; break; }
    }
    if (*sec_line < 0 || !key) return -1;
    for (int i = *sec_line + 1; i < *end; i++)
        if (key_is(L->line[i], key)) return i;
    return -1;
}

/* the value of a Key=Value line, without the blanks around it (and the quotes, as Windows does) */
static void value_of(const char *line, char *out, int cap)
{
    const char *v = strchr(line, '=');
    v = v ? v + 1 : "";
    while (*v == ' ' || *v == '\t') v++;
    int n = (int)strlen(v);
    while (n && (v[n - 1] == ' ' || v[n - 1] == '\t')) n--;
    if (n >= 2 && ((v[0] == '"' && v[n - 1] == '"') || (v[0] == '\'' && v[n - 1] == '\''))) { v++; n -= 2; }
    if (n > cap - 1) n = cap - 1;
    memcpy(out, v, (size_t)n); out[n] = 0;
}

/* ---------------------------------------------------------------- the profile API of Windows
 * GetPrivateProfileString, WritePrivateProfileString ... (dll.c, on Linux): .ini files of text */
int vb_ini_read(const char *path, const char *sec, const char *key, char *out, int cap)
{
    lines_t L; load(path, &L);
    int n = -1;
    if (cap <= 0) { free_lines(&L); return -1; }
    if (!sec || !key) {                         /* the sections, or the keys of a section: "a\0b\0\0" */
        int k = 0, sl, end;
        if (sec) find(&L, sec, 0, &sl, &end); else { sl = -1; end = L.n; }
        for (int i = sl + 1; i < end; i++) {
            const char *line = L.line[i], *e;
            char name[512]; int m = 0;
            if (!sec) {
                if (line[0] != '[' || !(e = strchr(line, ']'))) continue;
                m = (int)(e - line - 1); if (m > 500) m = 500; memcpy(name, line + 1, (size_t)m);
            } else {
                while (*line == ' ' || *line == '\t') line++;
                if (!*line || *line == ';' || !(e = strchr(line, '='))) continue;
                while (e > line && (e[-1] == ' ' || e[-1] == '\t')) e--;
                m = (int)(e - line); if (m > 500) m = 500; memcpy(name, line, (size_t)m);
            }
            if (k + m + 2 > cap) break;
            memcpy(out + k, name, (size_t)m); k += m; out[k++] = 0;
        }
        out[k] = 0;
        if (sec && sl < 0) k = 0;
        n = k;
    } else {
        int sl, end, i = find(&L, sec, key, &sl, &end);
        if (i >= 0) { value_of(L.line[i], out, cap); n = (int)strlen(out); }
    }
    free_lines(&L);
    return n;
}

int vb_ini_write(const char *path, const char *sec, const char *key, const char *value)
{
    lines_t L; load(path, &L);
    int sl, end, i = find(&L, sec, key, &sl, &end);
    if (!key) {                                 /* the whole section goes */
        if (sl >= 0) for (int j = end - 1; j >= sl; j--) del_line(&L, j);
    } else if (!value) {                        /* the key goes */
        if (i >= 0) del_line(&L, i);
    } else {
        size_t n = strlen(key) + strlen(value) + 2;
        char *ln = (char *)malloc(n); snprintf(ln, n, "%s=%s", key, value);
        if (i >= 0) { free(L.line[i]); L.line[i] = strdup(ln); }
        else if (sl >= 0) {
            int at = end;
            while (at > sl + 1 && !L.line[at - 1][0]) at--;     /* before the blank lines of the end */
            add_line(&L, ln, at);
        } else {
            char *h = (char *)malloc(strlen(sec) + 3); sprintf(h, "[%s]", sec);
            if (L.n && L.line[L.n - 1][0]) add_line(&L, "", -1);
            add_line(&L, h, -1); add_line(&L, ln, -1); free(h);
        }
        free(ln);
    }
    FILE *f = vbos_fopen(path, "wb");
    if (!f) { free_lines(&L); return 0; }
    for (int j = 0; j < L.n; j++) { fputs(L.line[j], f); fputs("\r\n", f); }
    fclose(f);
    free_lines(&L);
    return 1;
}

static char *dup8(vbs s) { return strdup(vb_stoutf8(s)); }

vbs vb_getsetting(vbs app, vbs section, vbs key, vbs def)
{
    char *a = dup8(app), *s = dup8(section), *k = dup8(key);
    lines_t L; load(file_of(a, 0), &L);
    int sl, end, i = find(&L, s, k, &sl, &end);
    char v[4096];
    if (i >= 0) value_of(L.line[i], v, sizeof v);
    vbs r = i >= 0 ? vb_sfromutf8(v) : vb_scopy(def);
    free_lines(&L); free(a); free(s); free(k);
    return r;
}

void vb_savesetting(vbs app, vbs section, vbs key, vbs value)
{
    char *a = dup8(app), *s = dup8(section), *k = dup8(key), *v = dup8(value);
    if (!*a || !*s || !*k) { vb_raise(5); goto out; }
    lines_t L; load(file_of(a, 1), &L);
    int sl, end, i = find(&L, s, k, &sl, &end);
    size_t n = strlen(k) + strlen(v) + 2;
    char *ln = (char *)malloc(n); snprintf(ln, n, "%s=%s", k, v);
    if (i >= 0) { free(L.line[i]); L.line[i] = strdup(ln); }
    else if (sl >= 0) add_line(&L, ln, end);
    else {
        char *h = (char *)malloc(strlen(s) + 3); sprintf(h, "[%s]", s);
        add_line(&L, h, -1); add_line(&L, ln, -1); free(h);
    }
    free(ln);
    save(file_of(a, 1), &L);
    free_lines(&L);
out:
    free(a); free(s); free(k); free(v);
}

/* DeleteSetting app, section [, key] */
void vb_deletesetting(vbs app, vbs section, vbs key, vb_i32 has_key)
{
    if (has_key < 0) {                              /* DeleteSetting app: all its sections */
        char *a = dup8(app);
        const char *f = file_of(a, 0);
        if (!*a || remove(f) != 0) vb_raise(5);
        free(a);
        return;
    }
    char *a = dup8(app), *s = dup8(section), *k = has_key ? dup8(key) : 0;
    lines_t L; load(file_of(a, 0), &L);
    int sl, end, i = find(&L, s, k, &sl, &end);
    if (sl < 0 || (k && i < 0)) vb_raise(5);         /* VB6: "Invalid procedure call" when it is not there */
    else {
        if (k) del_line(&L, i);
        else for (int j = end - 1; j >= sl; j--) del_line(&L, j);
        save(file_of(a, 0), &L);
    }
    free_lines(&L); free(a); free(s); free(k);
}

/* GetAllSettings(app, section): a Variant holding (0 To n-1, 0 To 1) keys and values; Empty when
 * the section is not there */
vbvar vb_getallsettings(vbs app, vbs section)
{
    char *a = dup8(app), *s = dup8(section);
    lines_t L; load(file_of(a, 0), &L);
    int sl, end;
    find(&L, s, 0, &sl, &end);
    vbvar v = vb_v_empty();
    int n = 0;
    if (sl >= 0)
        for (int i = sl + 1; i < end; i++) if (strchr(L.line[i], '=')) n++;
    if (n) {
        vbarr *arr = vb_atmp();
        vb_adim(arr, sizeof(vbvar), VBA_VAR, 0, 2, 0, n - 1, 0, 1);
        int k = 0;
        for (int i = sl + 1; i < end && arr->data; i++) {
            char *eq = strchr(L.line[i], '=');
            if (!eq) continue;
            *eq = 0;
            vb_vset((vbvar *)vb_aelem(arr, 2, k, 0), vb_v_s(vb_sfromutf8(L.line[i])));
            vb_vset((vbvar *)vb_aelem(arr, 2, k, 1), vb_v_s(vb_sfromutf8(eq + 1)));
            k++;
        }
        v.vt = VT_ARRAY | VT_VARIANT; v.u.a = arr;
    }
    free_lines(&L); free(a); free(s);
    return v;
}
