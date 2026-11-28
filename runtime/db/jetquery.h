/* jetquery.h - a saved query of Jet both ways: its rows in MSysQueries written as the SQL DAO shows, and Jet SQL
 * compiled into those rows (jetquery.c).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef JETQUERY_H
#define JETQUERY_H
#include "jetfile.h"                           /* (strndup on Windows) */

/* one row of MSysQueries; the texts in UTF-8 (0: Null), flag -32768: Null */
typedef struct jqrow { int attr; unsigned order; int flag; char *name1, *name2, *expr; } jqrow;
#define JQ_NOFLAG (-32768)

/* the SQL of a query from its rows (malloc); *qtype: the QueryDef's Type (0 select, 16 crosstab, 32 delete,
 * 48 update, 64 append, 80 make-table, 96 DDL, 128 union) */
char *jq_text(const jqrow *rows, int n, int *qtype);
/* Jet SQL compiled into rows (malloc'd array, *n of them); 0 and the message on a syntax error */
jqrow *jq_compile(const char *sql, int *n, char *err, int cap);
void jq_free(jqrow *rows, int n);
int jq_flags(const jqrow *rows, int n);        /* MSysObjects.Flags of the query */
#endif
