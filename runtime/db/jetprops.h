/* jetprops.h - the properties of MSysObjects.LvProp (jetprops.c).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef JETPROPS_H
#define JETPROPS_H
/* the end of a text that ends in a NUL (Modified UTF-8 of U+0000) */
#define JET_NUL "\xC0\x80"
#include "jetfile.h"

typedef struct jprop { char *name; int dtype, flags; unsigned char *v; int n; } jprop;   /* v: the bytes */
typedef struct jpblock { int kind; char *name; jprop *props; int nprops; } jpblock;     /* 0 object, 1 column, 2 index */
typedef struct jprops { int jet4, codepage; char **names; int nnames; jpblock *blocks; int nblocks; } jprops;

jprops *jprops_parse(const unsigned char *b, int n, int jet4, int codepage);   /* b 0: empty */
unsigned char *jprops_build(jprops *p, int *n);                                  /* malloc */
void jprops_free(jprops *p);
jpblock *jprops_block(jprops *p, int kind, const char *name, int make);
jprop *jprops_add(jpblock *b);
jprop *jprops_get(jpblock *b, const char *name);
void jprops_remove(jpblock *b, const char *name);
void jprops_set(jprops *p, jpblock *b, const char *name, int dtype, const void *v, int n, int flags);
void jprops_set_text(jprops *p, jpblock *b, const char *name, const char *s, int flags);
void jprops_set_bool(jprops *p, jpblock *b, const char *name, int v, int flags);
char *jprops_text(jprops *p, jprop *pr);                                         /* malloc, UTF-8 */
#endif
