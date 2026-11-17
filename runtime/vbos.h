/* vbos.h - the file system and the process for Linux and Windows (os.c); paths are UTF-8.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef VBOS_H
#define VBOS_H
#include <stdio.h>
#include <ctype.h>

FILE *vbos_fopen(const char *path, const char *mode);
int vbos_remove(const char *path);
int vbos_rename(const char *from, const char *to);
int vbos_mkdir(const char *path);
int vbos_rmdir(const char *path);
int vbos_chdir(const char *path);
void vbos_getcwd(char *out, int cap);
int vbos_stat(const char *path, long long *size, long long *mtime, int *attr);
int vbos_realdir(const char *path);              /* 1: a folder itself, not a symbolic link or junction to one */
void vbos_ci_path(char *p);                      /* Linux: the names of p as they are on the disk, without case */
void vbos_exe_path(char *out, int cap);
typedef struct vbos_dir vbos_dir;
vbos_dir *vbos_opendir(const char *path);
const char *vbos_readdir(vbos_dir *d, int *is_dir);
void vbos_closedir(vbos_dir *d);
int vbos_wildmatch(const char *pat, const char *name);
int vbos_copyfile(const char *from, const char *to);
long vbos_shell(const char *cmd, int style);
const char *vbos_getenv(const char *name);
const char *vbos_environ_n(int n);
void *vbos_dlopen(const char *lib);
void *vbos_dlsym(void *h, const char *name);
int vbos_ansi_codepage(void);
int vbos_user_lcid(void);
void vbos_console_utf8(void);
/* the Scripting runtime (scripting.c) */
int vbos_times(const char *path, long long *created, long long *accessed, long long *modified);
int vbos_disk(const char *path, double *total, double *freeb, double *avail);   /* bytes; 0 = ok */
int vbos_drive_info(const char *root, char *volume, int vcap, char *fs, int fcap, unsigned *serial, int *type);
int vbos_drives(char *out, int cap);             /* the roots, "\0" between them, "\0\0" at the end */
void vbos_special_folder(int which, char *out, int cap);   /* 0 Windows, 1 System, 2 Temp */
void vbos_abspath(const char *path, char *out, int cap);
int vbos_setattr(const char *path, int attr);
#endif
