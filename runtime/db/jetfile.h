/* jetfile.h - a Jet 3 / Jet 4 database (.mdb) as a model: read every table of a file (the system tables too, with
 * their long values) and write a model as a new file (page 0, the global usage map, TDEFs, data pages, long values,
 * indexes, usage maps), and the keys of Jet's indexes (General sort order). The format as mdbtools HACKING.md
 * describes it and as measured on the oracles of DEV-14 (VB6 + DAO 3.6); the Python prototype is
 * tools/jet/jetdb.py, jetsort.py, jetprops.py, jetquery.py.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef JETFILE_H
#define JETFILE_H
#include <stdint.h>
#include <stddef.h>
#if defined(_WIN32) && !defined(LB_STRNDUP)     /* (mingw-w64 has no strndup) */
#define LB_STRNDUP
#include <stdlib.h>
#include <string.h>
static inline char *lb_strndup(const char *s, size_t n)
{
    size_t k = 0;
    while (k < n && s[k]) k++;
    char *r = malloc(k + 1);
    if (r) { memcpy(r, s, k); r[k] = 0; }
    return r;
}
#define strndup lb_strndup
#endif

enum { JT_BOOL = 1, JT_BYTE = 2, JT_INT = 3, JT_LONG = 4, JT_MONEY = 5, JT_FLOAT = 6, JT_DOUBLE = 7, JT_DATETIME = 8,
       JT_BINARY = 9, JT_TEXT = 10, JT_OLE = 11, JT_MEMO = 12, JT_GUID = 15, JT_NUMERIC = 16 };

/* a value: the bytes as the row keeps them (fixed types), the text (Jet 4: UTF-16LE; Jet 3: the bytes of the
 * file's code page), the data of a long value; a Yes/No: n = 0 / 1. null: a Null (never for a Yes/No) */
typedef struct jv { unsigned char *p; int n; int null; } jv;

typedef struct jetcol {                   /* a column: its descriptor in the TDEF (raw: what is not understood) */
    char *name;                         /* UTF-8 */
    int type, num, len, flags, id, offv, offf, xflags, misc, misc2, fileflag;
    unsigned char raw[25]; int rawlen;
} jetcol;
typedef struct jridx {                  /* a real index: up to 10 (column number, order 1 ascending / 0 descending) */
    int n; int col[10]; unsigned char order[10];
    int flags, complex;
} jridx;
typedef struct jtable jtable;
typedef struct jlidx {                  /* a logical index: primary 1 / foreign key 2, its real index */
    char *name;
    int num, real, type, reltype, cup, cdel;
    uint32_t relidx, reltable;
    jtable *rel_to;                     /* a foreign key: the other table (its page is known when written) */
} jlidx;
struct jtable {
    char *name;
    uint32_t page;                      /* its TDEF page (read / written) */
    int type;                           /* 0x4e user, 0x53 system */
    int maxcols, nvar, keep_layout;     /* keep_layout: the offsets of the columns as read */
    jetcol *cols; int ncols;
    jridx *ridx; int nridx;
    jlidx *lidx; int nlidx;
    jv *vals; int nrows, caprows, width; /* the value of row r, column number c: vals[r * width + c] */
    uint32_t autonum, ct_autonum; int autonum_flag;
    int *lvorder; int nlv;              /* the long value columns in the order of the tail of the TDEF */
    void *user;
    int (*load)(jtable *t, void *ctx, char *err, int cap);  /* its rows, loaded by the writer when it writes it */
    void *loadctx;
    uint32_t usedmap;                   /* (read) the row of its usage map of the pages it uses: page << 8 | row */
};
typedef struct jdb {
    int jet4, ps;
    unsigned char *page0;               /* page 0, its masked part unmasked */
    jtable **t; int nt;                 /* MSysObjects first, then the other three catalog tables */
    int codepage;                       /* Jet 3: the code page of the texts (1252 ...) */
} jdb;

jdb *jet_read(const char *path, int codepage, char *err, int cap);
/* JET_NO_USER_ROWS: the rows of the user's tables not read (the export replaces them: memory) */
#define JET_NO_USER_ROWS 1
jdb *jet_read_opt(const char *path, int codepage, char *err, int cap, int opts);
int jet_write(jdb *db, const char *path, char *err, int cap);  /* temp + rename; 0 ok */
void jet_free(jdb *db);
jtable *jet_table(jdb *db, const char *name);
jetcol *jet_col(jtable *t, const char *name);
jv *jet_val(jtable *t, int row, int colnum);
jtable *jet_table_new(const char *name);
void jet_table_free(jtable *t);
int jet_row_add(jtable *t);                         /* a new row of Nulls (Yes/No: False); its number */
void jet_val_set(jv *v, const void *p, int n);       /* a copy of the bytes (p 0: Null) */

/* the keys of an index (jetsort.c) */
typedef struct jkey { unsigned char *p; int n, cap; } jkey;
void jet_sort_lcid(int lcid);                   /* Jet 4: the sort order of the keys made next (0x409 General) */
void jet_text_key(jkey *k, const unsigned char *text, int n, int jet4);  /* Jet 4: UTF-16LE; Jet 3: bytes */
void jet_col_key(jkey *k, int type, const jv *v, int desc, int jet4);
void jet_key_truncate(jkey *k, int jet4);

/* text: the code pages (the runtime's intl.c) and UTF-8 */
int jet_cp_to_utf16(int cp, const unsigned char *b, int n, uint16_t *out);   /* out: n units at most */
int jet_utf16_to_cp(int cp, const uint16_t *w, int n, unsigned char *out);   /* out: 2n bytes at most */
char *jet_utf16_to_utf8(const uint16_t *w, int n);  /* malloc */
int jet_utf8_to_utf16(const char *s, uint16_t **out); /* malloc; the count of units */
int jet_strcasecmp8(const char *a, const char *b);   /* names: case-insensitive (ASCII and Latin) */
/* jetnew.c: the creation date of page 0 of a new Jet 4 file and the key of its SIDs (measured together) */
void jet_new_datekey(unsigned char date[8], unsigned char key[2]);
#endif
