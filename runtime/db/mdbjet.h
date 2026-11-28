/* mdbjet.h - what mdbimport.c and mdbexport.c share: a Jet file read into SQLite and written back from it.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef MDBJET_H
#define MDBJET_H
#include "daoint.h"
#include "jetfile.h"
#include "jetprops.h"
#include "jetquery.h"

int mdb_codepage(void);                                     /* the code page of the texts of a Jet 3 file */
int jet_dao_type(int jt);                                   /* JT_ -> DataTypeEnum */
char *jet_text_utf8(jdb *db, const unsigned char *p, int n); /* a text of the file as UTF-8 (malloc) */
void jet_guid_text(const unsigned char *g, char out[40]);   /* {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx} */
double jet_numeric_value(const jetcol *c, const unsigned char *b);
jqrow *jet_query_rows(jdb *db, uint32_t oid, int *n);       /* the rows of a query in MSysQueries */
/* the properties of LvProp a program sees in Properties as its own (MSysLB_Properties): not those DAO has as
 * members (kind 'T' table, 'F' field, 'Q' query) and not those of binary types */
int mdb_builtin_prop(char kind, const char *name);
int mdb_prop_shown(int dtype);
char *mdb_prop_value_text(jprops *p, jprop *pr);            /* the value as MSysLB_Properties keeps it (malloc) */
jdb *jet_new(int jet4, int codepage);
jdb *jet_new_order(int jet4, int codepage, int lcid);   /* Jet 4 in another sort order (LCID) */
void jet_sid(int jet4, unsigned v, unsigned char out[2]);
/* the SQLite copy written as the Jet file path (made new when there is none: version 3 / 4 from MSysLB_Props) */
int mdb_export(sqlite3 *s, const char *path, char *err, int cap);
/* the password of a database: in page 0 (unmasked), or of a file (-1: not a Jet file) */
void jet_password_get(const unsigned char *p0, int jet4, int codepage, char *out, int cap);
int jet_password_set(unsigned char *p0, int jet4, int codepage, const char *pw);   /* -1: more than 20 characters */
int mdb_password(const char *path, char *out, int cap);
void jet_rc4(uint32_t key, unsigned char *data, int n);
#endif
