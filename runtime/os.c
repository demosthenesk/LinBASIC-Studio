/* os.c - the file system and the process, the same on Linux and Windows.
 * Paths are UTF-8 strings everywhere in the runtime; on Windows they become UTF-16 for the
 * wide (W) functions, so that names with Greek or other letters work in every code page.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
/* (no vbrt.h here: <windows.h> has its own VT_ names) */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "vbos.h"
#include <errno.h>
#include <sys/stat.h>
#ifdef _WIN32
#include <windows.h>
#include <direct.h>
#include <io.h>
#include <wchar.h>
#else
#include <unistd.h>
#include <dirent.h>
#include <sys/wait.h>
#include <sys/statvfs.h>
#include <limits.h>
#include <strings.h>
#endif

#ifdef _WIN32
static wchar_t *wide(const char *u8)               /* malloc'd UTF-16 */
{
    int n = MultiByteToWideChar(CP_UTF8, 0, u8, -1, 0, 0);
    wchar_t *w = (wchar_t *)malloc(sizeof(wchar_t) * (n > 0 ? n : 1));
    if (n > 0) MultiByteToWideChar(CP_UTF8, 0, u8, -1, w, n); else w[0] = 0;
    return w;
}
static void narrow(const wchar_t *w, char *out, int cap)
{
    int n = WideCharToMultiByte(CP_UTF8, 0, w, -1, out, cap, 0, 0);
    if (n <= 0) out[0] = 0;
}
#endif

FILE *vbos_fopen(const char *path, const char *mode)
{
#ifdef _WIN32
    wchar_t *p = wide(path), *m = wide(mode);
    FILE *f = _wfopen(p, m);
    free(p); free(m);
    return f;
#else
    return fopen(path, mode);
#endif
}

int vbos_remove(const char *path)
{
#ifdef _WIN32
    wchar_t *p = wide(path); int r = _wremove(p); free(p); return r;
#else
    return remove(path);
#endif
}

int vbos_rename(const char *from, const char *to)
{
#ifdef _WIN32
    wchar_t *a = wide(from), *b = wide(to); int r = _wrename(a, b); free(a); free(b); return r;
#else
    return rename(from, to);
#endif
}

int vbos_mkdir(const char *path)
{
#ifdef _WIN32
    wchar_t *p = wide(path); int r = _wmkdir(p); free(p); return r;
#else
    return mkdir(path, 0777);
#endif
}

int vbos_rmdir(const char *path)
{
#ifdef _WIN32
    wchar_t *p = wide(path); int r = _wrmdir(p); free(p); return r;
#else
    return rmdir(path);
#endif
}

int vbos_chdir(const char *path)
{
#ifdef _WIN32
    wchar_t *p = wide(path); int r = _wchdir(p); free(p); return r;
#else
    return chdir(path);
#endif
}

void vbos_getcwd(char *out, int cap)
{
#ifdef _WIN32
    wchar_t b[4096];
    if (!_wgetcwd(b, 4096)) b[0] = 0;
    narrow(b, out, cap);
#else
    if (!getcwd(out, (size_t)cap)) out[0] = 0;
#endif
}

/* size, modification time (seconds since 1970), attributes (vbDirectory 16, vbReadOnly 1, vbHidden 2) */
int vbos_stat(const char *path, long long *size, long long *mtime, int *attr)
{
#ifdef _WIN32
    wchar_t *p = wide(path);
    WIN32_FILE_ATTRIBUTE_DATA d;
    BOOL ok = GetFileAttributesExW(p, GetFileExInfoStandard, &d);
    free(p);
    if (!ok) return -1;
    if (size) *size = ((long long)d.nFileSizeHigh << 32) | d.nFileSizeLow;
    if (mtime) {
        unsigned long long t = ((unsigned long long)d.ftLastWriteTime.dwHighDateTime << 32) | d.ftLastWriteTime.dwLowDateTime;
        *mtime = (long long)(t / 10000000ULL) - 11644473600LL;
    }
    if (attr) *attr = (int)(d.dwFileAttributes & 0x37);
    return 0;
#else
    struct stat st;
    if (stat(path, &st) != 0) return -1;
    if (size) *size = (long long)st.st_size;
    if (mtime) *mtime = (long long)st.st_mtime;
    if (attr) {
        const char *b = strrchr(path, '/');
        b = b ? b + 1 : path;
        *attr = (S_ISDIR(st.st_mode) ? 16 : 0) | (access(path, W_OK) != 0 || !(st.st_mode & S_IWUSR) ? 1 : 0) |
                (b[0] == '.' && strcmp(b, ".") && strcmp(b, "..") ? 2 : 0);
    }
    return 0;
#endif
}

/* a folder itself (1), not a symbolic link / junction to one: a recursive delete goes into it, and only
 * removes the link of the other (rm -r does the same; through /proc/<pid>/root a link leads back to /) */
int vbos_realdir(const char *path)
{
#ifdef _WIN32
    wchar_t *p = wide(path);
    DWORD a = GetFileAttributesW(p);
    free(p);
    return a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_DIRECTORY) && !(a & FILE_ATTRIBUTE_REPARSE_POINT);
#else
    struct stat st;
    return lstat(path, &st) == 0 && S_ISDIR(st.st_mode);
#endif
}

/* a path of a VB program on Linux: a name that differs only in case from the one on the disk is that one (the
 * programs come from Windows, where case does not matter: Open "Jumble.skn" for JUMBLE.SKN - robot 6.8); a part that
 * does not exist stays as it is (a new file). Only the case changes, so p keeps its length */
void vbos_ci_path(char *p)
{
#ifndef _WIN32
    struct stat st;
    if (!p || !*p || stat(p, &st) == 0) return;
    size_t len = strlen(p);
    char *out = (char *)malloc(len + 2);
    if (!out) return;
    size_t o = 0;
    const char *s = p;
    if (*s == '/') { out[o++] = '/'; while (*s == '/') s++; }
    while (*s) {
        const char *e = strchr(s, '/');
        size_t n = e ? (size_t)(e - s) : strlen(s);
        memcpy(out + o, s, n); out[o + n] = 0;
        if (stat(out, &st) != 0 && !(n == 1 && *s == '.') && !(n == 2 && s[0] == '.' && s[1] == '.')) {
            char save = out[o]; out[o] = 0;
            DIR *d = opendir(o ? out : ".");
            out[o] = save;
            if (d) {
                struct dirent *de;
                while ((de = readdir(d)))
                    if (strlen(de->d_name) == n && !strncasecmp(de->d_name, s, n)) { memcpy(out + o, de->d_name, n); break; }
                closedir(d);
            }
        }
        o += n; s += n;
        while (*s == '/') { if (out[o - 1] != '/') out[o++] = '/'; s++; }
        out[o] = 0;
    }
    if (o == len) memcpy(p, out, len);
    free(out);
#else
    (void)p;
#endif
}

/* the executable of the program (App.Path, App.EXEName) */
void vbos_exe_path(char *out, int cap)
{
#ifdef _WIN32
    wchar_t b[4096];
    DWORD n = GetModuleFileNameW(0, b, 4096);
    if (n == 0 || n >= 4096) { out[0] = 0; return; }
    narrow(b, out, cap);
#else
    ssize_t n = readlink("/proc/self/exe", out, (size_t)cap - 1);
    out[n > 0 ? n : 0] = 0;
#endif
}

/* the names in a folder, one by one ("." and ".." left out) */
struct vbos_dir {
#ifdef _WIN32
    HANDLE h; WIN32_FIND_DATAW fd; int first;
#else
    DIR *d;
#endif
    char name[1024];
};

vbos_dir *vbos_opendir(const char *path)
{
    vbos_dir *d = (vbos_dir *)calloc(1, sizeof *d);
#ifdef _WIN32
    size_t n = strlen(path);
    char *pat = (char *)malloc(n + 3);
    snprintf(pat, n + 3, "%s%s*", path, n && (path[n - 1] == '\\' || path[n - 1] == '/') ? "" : "\\");
    wchar_t *w = wide(pat);
    d->h = FindFirstFileW(w, &d->fd);
    free(w); free(pat);
    if (d->h == INVALID_HANDLE_VALUE) { free(d); return 0; }
    d->first = 1;
#else
    d->d = opendir(path);
    if (!d->d) { free(d); return 0; }
#endif
    return d;
}

const char *vbos_readdir(vbos_dir *d, int *is_dir)
{
    for (;;) {
#ifdef _WIN32
        if (!d->first && !FindNextFileW(d->h, &d->fd)) return 0;
        d->first = 0;
        narrow(d->fd.cFileName, d->name, sizeof d->name);
        if (is_dir) *is_dir = (d->fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
#else
        struct dirent *e = readdir(d->d);
        if (!e) return 0;
        snprintf(d->name, sizeof d->name, "%s", e->d_name);
        if (is_dir) *is_dir = e->d_type == DT_DIR;
#endif
        if (strcmp(d->name, ".") && strcmp(d->name, "..")) return d->name;
    }
}

void vbos_closedir(vbos_dir *d)
{
    if (!d) return;
#ifdef _WIN32
    FindClose(d->h);
#else
    closedir(d->d);
#endif
    free(d);
}

/* * and ? of Dir, Kill ..., Like-free and without case (VB6 on Windows does not care about case); as Windows, a
 * name without an extension matches a pattern that ends in "." or ".*" ("*.*" is every file, "x.*" finds "x") */
static int wildmatch(const char *pat, const char *name, int nodot)
{
    if (*pat == 0) return *name == 0;
    if (*pat == '*') {
        while (*pat == '*') pat++;
        if (!*pat) return 1;
        for (;; name++) {
            if (wildmatch(pat, name, nodot)) return 1;
            if (!*name) return 0;
        }
    }
    if (!*name) return nodot && pat[0] == '.' && (pat[1] == 0 || (pat[1] == '*' && pat[2] == 0));
    if (*pat == '?' || tolower((unsigned char)*pat) == tolower((unsigned char)*name)) return wildmatch(pat + 1, name + 1, nodot);
    return 0;
}
int vbos_wildmatch(const char *pat, const char *name) { return wildmatch(pat, name, strchr(name, '.') == NULL); }

int vbos_copyfile(const char *from, const char *to)
{
#ifdef _WIN32
    wchar_t *a = wide(from), *b = wide(to);
    BOOL ok = CopyFileW(a, b, FALSE);
    free(a); free(b);
    return ok ? 0 : -1;
#else
    FILE *in = fopen(from, "rb");
    if (!in) return -1;
    FILE *out = fopen(to, "wb");
    if (!out) { fclose(in); return -2; }
    char buf[65536]; size_t n;
    while ((n = fread(buf, 1, sizeof buf, in)) > 0) if (fwrite(buf, 1, n, out) != n) { fclose(in); fclose(out); return -2; }
    fclose(in);
    return fclose(out) == 0 ? 0 : -2;
#endif
}

/* Shell: starts a program and does not wait; the task id (process id), or -1 */
#ifndef _WIN32
/* Shell of a program that does not exist: error 53 as VB6 (CreateProcess fails); "sh -c" alone always starts.
 * The program is the first word of the command (in quotes when it has spaces): a path, or found in PATH; a
 * builtin of the shell (cd, export ...) is let through. (aw24: Shell "gnome-system-monitor" without it) */
static int shell_program_exists(const char *cmd)
{
    while (*cmd == ' ' || *cmd == '\t') cmd++;
    char prog[1024]; size_t n = 0;
    if (*cmd == '"' || *cmd == '\'') {
        char q = *cmd++;
        while (*cmd && *cmd != q && n < sizeof prog - 1) prog[n++] = *cmd++;
    } else {
        while (*cmd && *cmd != ' ' && *cmd != '\t' && *cmd != ';' && *cmd != '&' && *cmd != '|' &&
               *cmd != '<' && *cmd != '>' && n < sizeof prog - 1) prog[n++] = *cmd++;
    }
    prog[n] = 0;
    if (!n) return 0;
    static const char *builtins[] = { "cd", "export", "exec", "set", "unset", "ulimit", "umask", ".", "source",
                                      "eval", "test", "[", "true", "false", "exit", "nohup", 0 };
    for (int i = 0; builtins[i]; i++) if (!strcmp(prog, builtins[i])) return 1;
    if (strchr(prog, '=')) return 1;            /* VAR=value program ... */
    if (strchr(prog, '/')) return access(prog, X_OK) == 0;
    const char *path = getenv("PATH");
    if (!path || !*path) path = "/usr/local/bin:/usr/bin:/bin";
    char full[2048];
    while (*path) {
        const char *e = strchr(path, ':');
        size_t len = e ? (size_t)(e - path) : strlen(path);
        snprintf(full, sizeof full, "%.*s/%s", (int)len, len ? path : ".", prog);
        struct stat stt;
        if (access(full, X_OK) == 0 && stat(full, &stt) == 0 && S_ISREG(stt.st_mode)) return 1;
        if (!e) break;
        path = e + 1;
    }
    return 0;
}
#endif

long vbos_shell(const char *cmd, int style)
{
#ifdef _WIN32
    wchar_t *w = wide(cmd);
    STARTUPINFOW si; PROCESS_INFORMATION pi;
    memset(&si, 0, sizeof si); si.cb = sizeof si;
    si.dwFlags = STARTF_USESHOWWINDOW;
    static const WORD show[] = { SW_HIDE, SW_SHOWNORMAL, SW_SHOWMINIMIZED, SW_SHOWMAXIMIZED, SW_SHOWNOACTIVATE, 0, SW_SHOWMINNOACTIVE };
    si.wShowWindow = style >= 0 && style <= 6 && style != 5 ? show[style] : SW_SHOWMINIMIZED;
    BOOL ok = CreateProcessW(0, w, 0, 0, FALSE, 0, 0, 0, &si, &pi);
    free(w);
    if (!ok) return -1;
    CloseHandle(pi.hThread); CloseHandle(pi.hProcess);
    return (long)pi.dwProcessId;
#else
    (void)style;
    if (!shell_program_exists(cmd)) return -1;
    fflush(0);
    pid_t pid = fork();
    if (pid < 0) return -1;
    if (pid == 0) {
        setsid();
        if (fork() == 0) { execl("/bin/sh", "sh", "-c", cmd, (char *)0); _exit(127); }
        _exit(0);
    }
    int st; waitpid(pid, &st, 0);               /* the grandchild runs on its own: no zombie */
    return (long)pid + 1;
#endif
}

/* Environ("NAME"), Environ(n): "NAME=value" of the n-th variable (1-based) */
const char *vbos_getenv(const char *name)
{
    return getenv(name);
}
const char *vbos_environ_n(int n)
{
#ifdef _WIN32
    extern char **_environ;
    char **e = _environ;
#else
    extern char **environ;
    char **e = environ;
#endif
    for (int i = 1; e && *e; e++, i++) if (i == n) return *e;
    return 0;
}

/* a DLL and one of its procedures (Declare on Windows) */
void *vbos_dlopen(const char *lib)
{
#ifdef _WIN32
    return (void *)LoadLibraryA(lib);
#else
    (void)lib;
    return 0;
#endif
}
void *vbos_dlsym(void *h, const char *name)
{
#ifdef _WIN32
    return (void *)GetProcAddress((HMODULE)h, name);
#else
    (void)h; (void)name;
    return 0;
#endif
}

/* Windows: the code page and the locale of the user; the console shows UTF-8 */
int vbos_ansi_codepage(void)
{
#ifdef _WIN32
    return (int)GetACP();
#else
    return 1252;
#endif
}
int vbos_user_lcid(void)
{
#ifdef _WIN32
    return (int)GetUserDefaultLCID();
#else
    return 1033;
#endif
}
void vbos_console_utf8(void)
{
#ifdef _WIN32
    SetConsoleOutputCP(65001);
#endif
}

/* ---------------------------------------------------------------- for the Scripting runtime */
/* created / accessed / modified, seconds since 1970 (Linux has no creation time: the change time) */
int vbos_times(const char *path, long long *created, long long *accessed, long long *modified)
{
#ifdef _WIN32
    wchar_t *p = wide(path);
    WIN32_FILE_ATTRIBUTE_DATA d;
    BOOL ok = GetFileAttributesExW(p, GetFileExInfoStandard, &d);
    free(p);
    if (!ok) return -1;
#define FT(x) ((long long)((((unsigned long long)(x).dwHighDateTime << 32) | (x).dwLowDateTime) / 10000000ULL) - 11644473600LL)
    if (created) *created = FT(d.ftCreationTime);
    if (accessed) *accessed = FT(d.ftLastAccessTime);
    if (modified) *modified = FT(d.ftLastWriteTime);
#undef FT
    return 0;
#else
    struct stat st;
    if (stat(path, &st) != 0) return -1;
    if (created) *created = (long long)(st.st_ctime < st.st_mtime ? st.st_ctime : st.st_mtime);
    if (accessed) *accessed = (long long)st.st_atime;
    if (modified) *modified = (long long)st.st_mtime;
    return 0;
#endif
}

int vbos_disk(const char *path, double *total, double *freeb, double *avail)
{
#ifdef _WIN32
    wchar_t *p = wide(path);
    ULARGE_INTEGER av, tot, fr;
    BOOL ok = GetDiskFreeSpaceExW(p, &av, &tot, &fr);
    free(p);
    if (!ok) return -1;
    if (total) *total = (double)tot.QuadPart;
    if (freeb) *freeb = (double)fr.QuadPart;
    if (avail) *avail = (double)av.QuadPart;
    return 0;
#else
    struct statvfs v;
    if (statvfs(path, &v) != 0) return -1;
    if (total) *total = (double)v.f_blocks * v.f_frsize;
    if (freeb) *freeb = (double)v.f_bfree * v.f_frsize;
    if (avail) *avail = (double)v.f_bavail * v.f_frsize;
    return 0;
#endif
}

/* volume name, file system, serial number, DriveType (1 removable, 2 fixed, 3 network, 4 CD-ROM, 5 RAM disk) */
int vbos_drive_info(const char *root, char *volume, int vcap, char *fs, int fcap, unsigned *serial, int *type)
{
#ifdef _WIN32
    wchar_t *r = wide(root), vol[MAX_PATH + 1], fsn[MAX_PATH + 1];
    DWORD ser = 0, maxc, fl;
    UINT t = GetDriveTypeW(r);
    BOOL ok = GetVolumeInformationW(r, vol, MAX_PATH + 1, &ser, &maxc, &fl, fsn, MAX_PATH + 1);
    free(r);
    if (!ok) { vol[0] = fsn[0] = 0; }
    if (volume) narrow(vol, volume, vcap);
    if (fs) narrow(fsn, fs, fcap);
    if (serial) *serial = ser;
    if (type) *type = t == DRIVE_REMOVABLE ? 1 : t == DRIVE_FIXED ? 2 : t == DRIVE_REMOTE ? 3 : t == DRIVE_CDROM ? 4 : t == DRIVE_RAMDISK ? 5 : 0;
    return ok ? 0 : -1;
#else
    if (volume && vcap > 0) volume[0] = 0;
    if (fs && fcap > 0) fs[0] = 0;
    if (serial) *serial = 0;
    if (type) *type = 2;
    FILE *m = fopen("/proc/mounts", "r");
    if (!m) return 0;
    char dev[1024], mnt[1024], typ[256];
    while (fscanf(m, "%1023s %1023s %255s %*[^\n]", dev, mnt, typ) == 3) {
        if (strcmp(mnt, root) != 0) continue;
        if (fs) snprintf(fs, (size_t)fcap, "%s", typ);
        if (type) {
            if (!strncmp(typ, "nfs", 3) || !strcmp(typ, "cifs") || !strcmp(typ, "smb3") || !strncmp(typ, "fuse.sshfs", 10)) *type = 3;
            else if (!strcmp(typ, "iso9660") || !strcmp(typ, "udf")) *type = 4;
            else if (!strcmp(typ, "tmpfs") || !strcmp(typ, "ramfs")) *type = 5;
            else if (!strncmp(mnt, "/media/", 7) || !strncmp(mnt, "/run/media/", 11)) *type = 1;
        }
        if (volume) {                           /* the label: the last part of a /media path */
            const char *b = strrchr(mnt, '/');
            snprintf(volume, (size_t)vcap, "%s", strcmp(mnt, "/") && b ? b + 1 : "");
        }
        if (serial) { unsigned h = 2166136261u; for (const char *q = dev; *q; q++) h = (h ^ (unsigned char)*q) * 16777619u; *serial = h; }
        break;
    }
    fclose(m);
    return 0;
#endif
}

int vbos_drives(char *out, int cap)
{
    int k = 0;
#ifdef _WIN32
    wchar_t b[1024];
    DWORD n = GetLogicalDriveStringsW(1023, b);
    for (wchar_t *p = b; n && *p; p += wcslen(p) + 1) {
        char d[16]; narrow(p, d, sizeof d);
        size_t l = strlen(d);
        if (l && d[l - 1] == '\\') d[--l] = 0;     /* "C:" */
        if (k + (int)l + 2 >= cap) break;
        memcpy(out + k, d, l + 1); k += (int)l + 1;
    }
#else
    const char *roots[64]; int nr = 0;
    static char mnts[64][1024];
    roots[nr++] = "/";
    FILE *m = fopen("/proc/mounts", "r");
    char dev[1024], mnt[1024];
    while (m && nr < 64 && fscanf(m, "%1023s %1023s %*[^\n]", dev, mnt) == 2)
        if (!strncmp(mnt, "/media/", 7) || !strncmp(mnt, "/mnt/", 5) || !strncmp(mnt, "/run/media/", 11)) {
            int dup = 0;
            for (int i = 0; i < nr; i++) if (!strcmp(roots[i], mnt)) dup = 1;
            if (!dup) { snprintf(mnts[nr], sizeof mnts[nr], "%s", mnt); roots[nr] = mnts[nr]; nr++; }
        }
    if (m) fclose(m);
    for (int i = 0; i < nr; i++) {
        int l = (int)strlen(roots[i]);
        if (k + l + 2 >= cap) break;
        memcpy(out + k, roots[i], (size_t)l + 1); k += l + 1;
    }
#endif
    out[k] = 0;
    return k;
}

void vbos_special_folder(int which, char *out, int cap)
{
#ifdef _WIN32
    wchar_t b[MAX_PATH + 1];
    UINT n = which == 0 ? GetWindowsDirectoryW(b, MAX_PATH) : which == 1 ? GetSystemDirectoryW(b, MAX_PATH) : GetTempPathW(MAX_PATH, b);
    if (!n) b[0] = 0;
    size_t l = wcslen(b);
    if (which == 2 && l > 3 && b[l - 1] == L'\\') b[l - 1] = 0;
    narrow(b, out, cap);
#else
    const char *t = getenv("TMPDIR");
    if (which == 2) { snprintf(out, (size_t)cap, "%s", t && *t ? t : "/tmp"); return; }
    /* the Windows and System folders of the Linux emulation (dll.c: GetWindowsDirectoryA) */
    const char *h = getenv("XDG_DATA_HOME"), *home = getenv("HOME");
    char base[3000];
    if (h && *h) snprintf(base, sizeof base, "%s/linbasic-studio/windows", h);
    else snprintf(base, sizeof base, "%s/.local/share/linbasic-studio/windows", home ? home : "/tmp");
    snprintf(out, (size_t)cap, which == 0 ? "%s" : "%s/system32", base);
#endif
}

void vbos_abspath(const char *path, char *out, int cap)
{
#ifdef _WIN32
    wchar_t *p = wide(path), b[4096];
    DWORD n = GetFullPathNameW(p, 4096, b, 0);
    free(p);
    if (!n || n >= 4096) { snprintf(out, (size_t)cap, "%s", path); return; }
    narrow(b, out, cap);
#else
    char cwd[4096], tmp[8192];
    if (path[0] == '/') snprintf(tmp, sizeof tmp, "%s", path);
    else { if (!getcwd(cwd, sizeof cwd)) cwd[0] = 0; snprintf(tmp, sizeof tmp, "%s/%s", cwd, path); }
    /* . and .. without touching the disk */
    char *parts[1024]; int np = 0;
    for (char *t = strtok(tmp, "/"); t && np < 1024; t = strtok(0, "/")) {
        if (!strcmp(t, ".")) continue;
        if (!strcmp(t, "..")) { if (np) np--; continue; }
        parts[np++] = t;
    }
    int k = 0;
    out[0] = 0;
    if (!np) { snprintf(out, (size_t)cap, "/"); return; }
    for (int i = 0; i < np && k < cap - 1; i++) k += snprintf(out + k, (size_t)(cap - k), "/%s", parts[i]);
#endif
}

/* Attributes =: vbReadOnly 1, vbHidden 2, vbSystem 4, vbArchive 32 (Linux: only read-only) */
int vbos_setattr(const char *path, int attr)
{
#ifdef _WIN32
    wchar_t *p = wide(path);
    DWORD a = GetFileAttributesW(p);
    if (a == INVALID_FILE_ATTRIBUTES) { free(p); return -1; }
    a = (a & ~(DWORD)0x27) | (DWORD)(attr & 0x27);
    BOOL ok = SetFileAttributesW(p, a);
    free(p);
    return ok ? 0 : -1;
#else
    struct stat st;
    if (stat(path, &st) != 0) return -1;
    mode_t m = st.st_mode & 07777;
    if (attr & 1) m &= ~(mode_t)0222; else m |= S_IWUSR;
    return chmod(path, m);
#endif
}
