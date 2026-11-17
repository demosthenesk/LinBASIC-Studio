/* vbrt_int.h - internal declarations shared by the libvbrt sources.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef VBRT_INT_H
#define VBRT_INT_H
#include "vbrt.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>
#include <math.h>
#ifdef _WIN32
#include <time.h>
#define localtime_r(t, tm) localtime_s((tm), (t))
#define gmtime_r(t, tm) gmtime_s((tm), (t))
#endif

extern int vb_codepage;                          /* 1252, 1253 ... */
/* intl.c: the code pages of Windows, case, StrConv, text comparison */
int vb_cp_decode(int cp, const unsigned char *b, int n, vb_wchar *out);
int vb_cp_encode(int cp, const vb_wchar *w, int n, unsigned char *out);
int vb_cp_code(int cp, vb_wchar u, int bestfit);
int vb_cp_isdbcs(int cp);
int vb_cp_islead(int cp, unsigned char c);
int vb_lcid_cp(vb_i32 lcid);
vb_wchar vb_wupper(vb_wchar c);
vb_wchar vb_wlower(vb_wchar c);
vb_i32 vb_collate(vbs a, vbs b);
unsigned vb_collate_hash(vbs s);                /* the same for the Strings vb_collate finds equal */
vb_wchar vb_ansi2w(unsigned char c);
int vb_w2ansi(vb_wchar w);                       /* -1 if not representable */
vbs vb_sfromansi(const char *s, int n);
char *vb_stoansi(vbs s, int *n);                 /* temporary buffer */
vbs vb_sfrombytes(const char *b, vb_i32 n);      /* two bytes in each character */
int vb_ini_read(const char *path, const char *sec, const char *key, char *out, int cap);   /* -1: not there */
int vb_ini_write(const char *path, const char *sec, const char *key, const char *value);
enum { VB_FMT_STR = 0, VB_FMT_LOCALE = 1 };
extern int vb_lcid, vb_lc_lzero;
extern char vb_lc_dec, vb_lc_grp;
extern const char *vb_lc_sdate, *vb_lc_ltime, *vb_lc_am, *vb_lc_pm;   /* UTF-8 */
extern const char *vb_lc_ldate, *vb_lc_stime, *vb_lc_cursym, *vb_lc_curpos, *vb_lc_curneg;
extern int vb_lc_names, vb_lc_firstday;
int vb_lc_seps(int lcid, char *dec, char *grp);   /* the separators of a language; 0: not in the table */
const char *vb_lc_month(int m, int kind);        /* kind 0 name, 1 abbreviation, 2 genitive (Greek) */
const char *vb_lc_day(int wd, int abbrev);         /* wd 0 = Sunday */
int vb_fmt_pic(char *o, int cap, const char *pic, double date);   /* Windows date/time picture */
void vb_fmt_real(double v, int digits, int mode, char *out);
void vb_fmt_cy(vb_cy c, int mode, char *out);
void vb_date_iso(vb_date v, char *out);
vbs vb_sascii(const char *s);
const char *vb_errtext(int n);
void vb_set_error_text(vbs *dst, int n);
void vb_files_flush(void);
typedef struct vbfile { FILE *f; int mode; int col; int reclen; int open; int eofflag; vb_i32 lastrec; } vbfile;
vbfile *vb_getfile(vb_i32 n, int need_mode);
#endif
