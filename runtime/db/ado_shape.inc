/* ado_shape.inc - the provider MSDataShape of ADO: hierarchical recordsets (Data Shaping), the
 * rows of the DataEnvironment of VB6 and of its DataReport. Included by ado.c.
 *   SHAPE {command} [AS alias] APPEND ({child} [AS alias] RELATE parentcol TO childcol [, ...]) [AS name],
 *         SUM|AVG|MIN|MAX|COUNT|ANY|STDEV(alias.field) [AS name], CALC(expression) [AS name], ...
 *   SHAPE {command} AS alias COMPUTE alias, aggregates ..., CALC(...) [BY field, ...]
 *   ... RELATE parentcol TO PARAMETER n: the child command again for each parent row
 * The rows come from the provider under it (Jet, ODBC) into recordsets in memory; a chapter field
 * (adChapter) gives a recordset of the child rows of the current parent row, which follows the
 * parent when it moves (StayInSync); changes go back to the server through the client cursor
 * (cursor_write of ado_odbc.inc). Types, attributes, names and errors as VB6 gives them (VB6 oracle
 * tests/db/d_shape1).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

#define SRC_SHAPE "MSDataShape"

/* a chapter column: where its rows come from */
typedef struct achap {
    int refs;
    vbobj child;                        /* the recordset of all the child rows (counted) */
    int n, pcol[16], ccol[16];          /* parent column = child column */
    char *psql;                         /* RELATE ... TO PARAMETER: the child command, again for each row */
    int pidx[16];                       /* the parameter each pair gives */
    int indexed;                        /* the child rows have an index on the relation */
} achap;
static void chap_free(achap *a)
{
    if (!a || --a->refs > 0) return;
    if (a->child) vb_orelease(a->child);
    free(a->psql);
    free(a);
}

/* ---------------------------------------------------------------- the command text */
enum { SF_REL = 1, SF_AGG, SF_CALC, SF_NEW, SF_CHAP, SF_FIELD };
struct shp;
typedef struct stab { char *cmd; struct shp *sub; int pos; } stab;   /* {provider command} or (SHAPE ...) */
typedef struct sfld {
    int kind;
    char *alias;
    stab tab; char *talias;             /* SF_REL: the child */
    int n; char *pn[16], *cn[16]; int par[16];  /* RELATE p TO c / TO PARAMETER n (par = n + 1) */
    char *fn;                           /* SF_AGG: SUM ...; SF_CALC: the expression; SF_NEW: the type */
    char *qal, *qfld;                   /* SF_AGG: alias.field (qfld 0: COUNT(alias)) */
    int newtype, newsize;
} sfld;
typedef struct shp {
    stab tab; char *alias;
    int compute;
    sfld *f; int nf;
    char **by; int nby;
} shp;

static void stab_free(stab *t);
static void shp_free(shp *s)
{
    if (!s) return;
    stab_free(&s->tab); free(s->alias);
    for (int i = 0; i < s->nf; i++) {
        sfld *f = &s->f[i];
        free(f->alias); stab_free(&f->tab); free(f->talias); free(f->fn); free(f->qal); free(f->qfld);
        for (int k = 0; k < f->n; k++) { free(f->pn[k]); free(f->cn[k]); }
    }
    free(s->f);
    for (int i = 0; i < s->nby; i++) free(s->by[i]);
    free(s->by);
    free(s);
}
static void stab_free(stab *t) { free(t->cmd); shp_free(t->sub); t->cmd = 0; t->sub = 0; }

typedef struct spar { const char *src, *p; int err; int errpos; } spar;
static void sp_ws(spar *s) { while (*s->p && isspace((unsigned char)*s->p)) s->p++; }
static int sp_fail(spar *s) { if (!s->err) { s->err = 1; s->errpos = (int)(s->p - s->src); } return -1; }
static int sp_word(spar *s, const char *w)        /* a keyword */
{
    sp_ws(s);
    size_t n = strlen(w);
    if (strncasecmp(s->p, w, n) || isalnum((unsigned char)s->p[n]) || s->p[n] == '_') return 0;
    s->p += n;
    return 1;
}
static int sp_ch(spar *s, char c) { sp_ws(s); if (*s->p == c) { s->p++; return 1; } return 0; }
static char *sp_name(spar *s)                     /* a name: plain, [bracketed], 'quoted', "quoted" or `quoted` */
{
    sp_ws(s);
    const char *p = s->p;
    jbuf b = { 0 };
    if (*p == '[' || *p == '\'' || *p == '"' || *p == '`') {
        char q = *p == '[' ? ']' : *p;
        p++;
        while (*p && *p != q) jb_put(&b, p++, 1);
        if (*p != q) { jb_free(&b); return 0; }
        p++;
    } else {
        while (*p && (isalnum((unsigned char)*p) || *p == '_' || (unsigned char)*p >= 0x80)) jb_put(&b, p++, 1);
        if (!b.n) return 0;
    }
    s->p = p;
    if (!b.p) jb_s(&b, "");
    return b.p;
}
static int is_kw(const char *n)
{
    static const char *const kw[] = { "AS", "APPEND", "COMPUTE", "BY", "RELATE", "TO", "PARAMETER", "SHAPE", "CALC", "NEW", "TABLE", 0 };
    for (int i = 0; kw[i]; i++) if (!strcasecmp(n, kw[i])) return 1;
    return 0;
}
static char *sp_alias(spar *s)                    /* [AS] alias, or 0 */
{
    const char *keep = s->p;
    int as = sp_word(s, "AS");
    const char *at = s->p;
    char *n = sp_name(s);
    if (n && !as && is_kw(n)) { free(n); s->p = keep; return 0; }
    if (!n) { if (as) { s->p = at; sp_fail(s); } else s->p = keep; return 0; }
    return n;
}
static shp *sp_shape(spar *s);
static int sp_table(spar *s, stab *t)             /* {command} | (SHAPE ...) | TABLE name | name */
{
    sp_ws(s);
    t->pos = (int)(s->p - s->src);
    if (*s->p == '{') {
        const char *st = ++s->p;
        int depth = 1;
        while (*s->p && depth) { if (*s->p == '{') depth++; else if (*s->p == '}') depth--; if (depth) s->p++; }
        if (!*s->p) return sp_fail(s);
        t->cmd = malloc((size_t)(s->p - st) + 1);
        memcpy(t->cmd, st, (size_t)(s->p - st)); t->cmd[s->p - st] = 0;
        s->p++;
        return 0;
    }
    if (*s->p == '(') {
        const char *keep = s->p;
        s->p++;
        if (sp_word(s, "SHAPE")) {
            s->p -= 5;
            t->sub = sp_shape(s);
            if (!t->sub) return -1;
            if (!sp_ch(s, ')')) return sp_fail(s);
            return 0;
        }
        s->p = keep;
        return sp_fail(s);
    }
    if (sp_word(s, "TABLE")) {}
    char *n = sp_name(s);
    if (!n) return sp_fail(s);
    jbuf b = { 0 }; jb_s(&b, "SELECT * FROM ["); jb_s(&b, n); jb_s(&b, "]");
    free(n);
    t->cmd = b.p;
    return 0;
}
/* alias.field */
static int sp_qual(spar *s, char **al, char **fl)
{
    char *a = sp_name(s);
    if (!a) return sp_fail(s);
    if (sp_ch(s, '.')) {
        char *f = sp_name(s);
        if (!f) { free(a); return sp_fail(s); }
        *al = a; *fl = f;
    } else { *al = a; *fl = 0; }
    return 0;
}
/* CALC( ... ): the expression up to its closing parenthesis */
static char *sp_paren_text(spar *s)
{
    if (!sp_ch(s, '(')) { sp_fail(s); return 0; }
    const char *st = s->p;
    int depth = 1; char q = 0;
    while (*s->p) {
        char c = *s->p;
        if (q) { if (c == q) q = 0; }
        else if (c == '"' || c == '\'' || c == '#') q = c == '#' ? '#' : c;
        else if (c == '(') depth++;
        else if (c == ')' && !--depth) break;
        s->p++;
    }
    if (!*s->p) { sp_fail(s); return 0; }
    char *r = malloc((size_t)(s->p - st) + 1);
    memcpy(r, st, (size_t)(s->p - st)); r[s->p - st] = 0;
    s->p++;
    return r;
}
static const struct { const char *name; int type, size; } NEWTYPES[] = {
    { "adInteger", 3, 4 }, { "adSmallInt", 2, 2 }, { "adBigInt", 20, 8 }, { "adTinyInt", 16, 1 }, { "adUnsignedTinyInt", 17, 1 },
    { "adSingle", 4, 4 }, { "adDouble", 5, 8 }, { "adCurrency", 6, 8 }, { "adDate", 7, 8 }, { "adBoolean", 11, 2 },
    { "adBSTR", 8, 255 }, { "adChar", 129, 255 }, { "adVarChar", 200, 255 }, { "adWChar", 130, 255 }, { "adVarWChar", 202, 255 },
    { "adNumeric", 131, 19 }, { "adDecimal", 14, 19 }, { "adDBDate", 133, 6 }, { "adDBTimeStamp", 135, 16 }, { "adVariant", 12, 16 },
    { 0, 0, 0 } };
static int sp_field(spar *s, sfld *f)
{
    sp_ws(s);
    if (*s->p == '(') {                         /* (child RELATE ...) */
        s->p++;
        f->kind = SF_REL;
        if (sp_table(s, &f->tab)) return -1;
        f->talias = sp_alias(s);
        if (s->err) return -1;
        if (!sp_word(s, "RELATE")) return sp_fail(s);
        do {
            if (f->n >= 16) return sp_fail(s);
            char *pn = sp_name(s);
            if (!pn) return sp_fail(s);
            f->pn[f->n] = pn;
            if (!sp_word(s, "TO")) { f->n++; return sp_fail(s); }
            if (sp_word(s, "PARAMETER")) {
                sp_ws(s);
                if (!isdigit((unsigned char)*s->p)) { f->n++; return sp_fail(s); }
                f->par[f->n] = (int)strtol(s->p, (char **)&s->p, 10) + 1;
            } else {
                char *cn = sp_name(s);
                if (!cn) { f->n++; return sp_fail(s); }
                f->cn[f->n] = cn;
            }
            f->n++;
        } while (sp_ch(s, ','));
        if (!sp_ch(s, ')')) return sp_fail(s);
    } else if (sp_word(s, "CALC")) {
        f->kind = SF_CALC;
        f->fn = sp_paren_text(s);
        if (!f->fn) return -1;
    } else if (sp_word(s, "NEW")) {
        f->kind = SF_NEW;
        char *t = sp_name(s);
        if (!t) return sp_fail(s);
        for (int i = 0; NEWTYPES[i].name; i++)
            if (!strcasecmp(t, NEWTYPES[i].name)) { f->newtype = NEWTYPES[i].type; f->newsize = NEWTYPES[i].size; }
        free(t);
        if (!f->newtype) return sp_fail(s);
        if (sp_ch(s, '(')) {
            sp_ws(s); f->newsize = (int)strtol(s->p, (char **)&s->p, 10);
            if (sp_ch(s, ',')) { sp_ws(s); strtol(s->p, (char **)&s->p, 10); }
            if (!sp_ch(s, ')')) return sp_fail(s);
        }
    } else {
        static const char *const AGG[] = { "SUM", "AVG", "MIN", "MAX", "COUNT", "ANY", "STDEV", 0 };
        for (int i = 0; AGG[i]; i++) {
            const char *keep = s->p;
            if (sp_word(s, AGG[i]) && sp_ch(s, '(')) {
                f->kind = SF_AGG; f->fn = strdup(AGG[i]);
                if (sp_qual(s, &f->qal, &f->qfld)) return -1;
                if (!f->qfld && strcmp(AGG[i], "COUNT")) return sp_fail(s);
                if (!sp_ch(s, ')')) return sp_fail(s);
                break;
            }
            s->p = keep;
        }
        if (!f->kind) {                         /* COMPUTE: the alias of the child, or a field */
            char *n = sp_name(s);
            if (!n || is_kw(n)) { free(n); return sp_fail(s); }
            f->kind = SF_FIELD; f->fn = n;
        }
    }
    f->alias = sp_alias(s);
    return s->err ? -1 : 0;
}
static shp *sp_shape(spar *s)
{
    if (!sp_word(s, "SHAPE")) { sp_fail(s); return 0; }
    shp *r = calloc(1, sizeof *r);
    sp_ws(s);
    if (*s->p && strncasecmp(s->p, "APPEND", 6) && strncasecmp(s->p, "COMPUTE", 7)) {
        if (sp_table(s, &r->tab)) { shp_free(r); return 0; }
        r->alias = sp_alias(s);
        if (s->err) { shp_free(r); return 0; }
    }
    int app = sp_word(s, "APPEND"), comp = !app && sp_word(s, "COMPUTE");
    if (app || comp) {
        r->compute = comp;
        do {
            r->f = realloc(r->f, sizeof(sfld) * (size_t)(r->nf + 1));
            memset(&r->f[r->nf], 0, sizeof(sfld));
            r->nf++;
            if (sp_field(s, &r->f[r->nf - 1])) { shp_free(r); return 0; }
        } while (sp_ch(s, ','));
        if (comp && sp_word(s, "BY")) {
            do {
                char *n = sp_name(s);
                if (!n) { sp_fail(s); shp_free(r); return 0; }
                r->by = realloc(r->by, sizeof(char *) * (size_t)(r->nby + 1));
                r->by[r->nby++] = n;
            } while (sp_ch(s, ','));
        }
    }
    if (!r->tab.cmd && !r->tab.sub && !app) { sp_fail(s); shp_free(r); return 0; }
    return r;
}
/* the error of MSDataShape for a text it cannot read */
static void shape_syntax_error(const char *text, int pos)
{
    int len = (int)strlen(text);
    if (pos > len) pos = len;
    int from = pos - 18 < 0 ? 0 : pos - 18;
    char near[64]; snprintf(near, sizeof near, "%.*s", pos - from, text + from);
    char msg[400];
    snprintf(msg, sizeof msg, "Data shape command text contains a syntax error at or near position %d in the command. The text near the error is: \"%s\".", pos, near);
    raise(-2147217900, SRC_SHAPE, msg);
}

/* ---------------------------------------------------------------- the rows of a provider command */
static void jet_target(ars *r, ddb *d, drs *b)
{
    ot_free(r->ot); r->ot = 0;
    if (!b->base) return;
    otarget *t = calloc(1, sizeof *t);
    t->table = strdup(b->base); t->ncols = b->ncols; t->autocol = -1;
    t->cols = calloc((size_t)(b->ncols ? b->ncols : 1), sizeof(char *));
    t->iskey = calloc((size_t)(b->ncols ? b->ncols : 1), sizeof(int));
    for (int i = 0; i < b->ncols; i++) {
        rcol *c = &b->cols[i];
        if (c->stable && strcasecmp(c->stable, b->base)) continue;
        if (c->sfield) t->cols[i] = strdup(c->sfield);
        else if (c->stable) t->cols[i] = strdup(c->name);
        if (c->autoinc && t->autocol < 0) t->autocol = i;
    }
    dtd *td = schema_load_table(d, b->base);
    if (td) {
        for (int k = 0; k < td->indexes->n; k++) {
            dix *x = (dix *)vb_oid(td->indexes->items[k]);
            if (!x->primary) continue;
            for (int j = 0; j < x->fields->n; j++) {
                dfld *f = (dfld *)vb_oid(x->fields->items[j]);
                for (int i = 0; i < b->ncols; i++) if (t->cols[i] && !strcasecmp(t->cols[i], f->name)) t->iskey[i] = 1;
            }
        }
        vb_orelease((vbobj)td);
    }
    r->ot = t;
}
static int cursor_write(void *ctx, drs *d, int op);
/* a command of the provider under MSDataShape: its rows in memory, its changes to go to the server */
static ars *shape_exec(acn *c, const char *text, vbvar *vals, int nv)
{
    ars *r = (ars *)rs_new_for(c);
    r->cursortype = adOpenStatic; r->cursorloc = adUseClient; r->locktype = adLockOptimistic;
    if (c->odbc) {
        int rows = 0;
        char *src; char *sql = odbc_sql(text, adCmdUnknown, 0, &src);
        int e = odbc_run(c, sql, vals, nv, r, 0, &rows);
        free(sql); free(src);
        if (e) { vb_orelease((vbobj)r); return 0; }
        if (!rows) { vb_orelease((vbobj)r); raise(-2147217900, SRC_SHAPE, "Command text did not return a rowset."); return 0; }
        r->state = 1;
        return r;
    }
    ddb *d = DB(c->db);
    char *sql, *src;
    if (resolve_cmd(c, d, text, adCmdUnknown, &sql, &src)) { vb_orelease((vbobj)r); return 0; }
    free(src);
    vbobj o = 0; long n;
    int e = run_sql(c, d, sql, vals, nv, dbOpenDynaset, &o, &n);   /* a dynaset knows its base table */
    if (e) {                                    /* not updatable (GROUP BY ...): a snapshot */
        vb_err_clear_pending(); alist_clear(c->errs);
        e = run_sql(c, d, sql, vals, nv, dbOpenSnapshot, &o, &n);
    }
    free(sql);
    if (e) { vb_orelease((vbobj)r); return 0; }
    if (!o) { vb_orelease((vbobj)r); raise(-2147217900, SRC_SHAPE, "Command text did not return a rowset."); return 0; }
    rs_from_drs(r, o, 0);
    vb_orelease(o);
    drs *b = (drs *)vb_oid(r->base);
    /* where the changes go: before the rows go to memory */
    otarget *keep_t = 0;
    { ars tmp; memset(&tmp, 0, sizeof tmp); jet_target(&tmp, d, b); keep_t = tmp.ot; }
    if (disconnect(r)) { ot_free(keep_t); vb_orelease((vbobj)r); return 0; }
    r->cn = (vbobj)c; vb_oaddref((vbobj)c);
    r->ot = keep_t;
    DB(r->db)->remote = cursor_write; DB(r->db)->remote_ctx = r;
    r->state = 1;
    return r;
}

/* ---------------------------------------------------------------- building the hierarchy */
static int col_named(ars *r, const char *name)
{
    for (int i = 0; i < r->ndefs; i++) if (!strcasecmp(r->defs[i].name, name)) return i;
    return -1;
}
static int veq(vbvar a, vbvar b)
{
    int an = a.vt == VT_NULL || a.vt == VT_EMPTY, bn = b.vt == VT_NULL || b.vt == VT_EMPTY;
    if (an || bn) return an && bn;
    vb_i32 save = vb_errpend;
    vbvar r = vb_v_cmp(a, b, 0, 0);
    if (vb_errpend && !save) { vb_err_clear_pending(); return 0; }
    return r.vt == VT_BOOL && r.u.b != 0;
}
static int vless(vbvar a, vbvar b)            /* the order of COMPUTE ... BY: Null first */
{
    int an = a.vt == VT_NULL || a.vt == VT_EMPTY, bn = b.vt == VT_NULL || b.vt == VT_EMPTY;
    if (an || bn) return an && !bn;
    vb_i32 save = vb_errpend;
    vbvar r = vb_v_cmp(a, b, 2, 1);
    if (vb_errpend && !save) { vb_err_clear_pending(); return 0; }
    return r.vt == VT_BOOL && r.u.b != 0;
}
static drs *BASE(ars *r) { return (drs *)vb_oid(r->base); }

/* an aggregate over the child rows that go with the values pv of the parent */
static vbvar aggregate(const char *fn, ars *ch, int fcol, achap *a, vbvar *pv)
{
    drs *b = BASE(ch);
    vbvar acc = vb_v_null(); int n = 0; double sum = 0, sum2 = 0;
    int cy = fcol >= 0 && ch->defs[fcol].type == adCurrency;
    for (int i = 0; i < b->d->n; i++) {
        if (b->d->r[i].del) continue;
        vbvar *row = b->d->r[i].v;
        int ok = 1;
        for (int k = 0; k < a->n && ok; k++) ok = veq(pv[k], row[a->ccol[k]]);
        if (!ok) continue;
        if (fcol < 0) { n++; continue; }        /* COUNT(alias) */
        vbvar x = row[fcol];
        if (x.vt == VT_NULL || x.vt == VT_EMPTY) continue;
        n++;
        if (!strcmp(fn, "SUM") || !strcmp(fn, "AVG") || !strcmp(fn, "STDEV")) {
            double d = vb_v2d(x); sum += d; sum2 += d * d;
            if (cy) acc = acc.vt == VT_NULL ? vb_v_c(vb_v2c(x)) : vb_v_add(acc, vb_v_c(vb_v2c(x)));
        } else if (!strcmp(fn, "MIN")) { if (acc.vt == VT_NULL || vless(x, acc)) vb_vset(&acc, x); }
        else if (!strcmp(fn, "MAX")) { if (acc.vt == VT_NULL || vless(acc, x)) vb_vset(&acc, x); }
        else if (!strcmp(fn, "ANY")) { if (acc.vt == VT_NULL) vb_vset(&acc, x); }
    }
    if (!strcmp(fn, "COUNT")) return vb_v_l(n);
    if (!n) return vb_v_null();
    if (!strcmp(fn, "SUM")) return cy ? acc : vb_v_d(sum);
    if (!strcmp(fn, "AVG")) return cy ? vb_v_div(acc, vb_v_l(n)) : vb_v_d(sum / n);
    if (!strcmp(fn, "STDEV")) return n > 1 ? vb_v_d(sqrt((sum2 - sum * sum / n) / (n - 1))) : vb_v_null();
    return acc;
}

static vbvar calc_eval(const char *expr, ars *r, vbvar *row, int ncols, int *err);

/* the recordset of a SHAPE command */
static ars *shape_eval(acn *c, shp *s, const char *text);
static ars *shape_tab(acn *c, stab *t, const char *text)
{
    if (t->sub) return shape_eval(c, t->sub, text);
    return shape_exec(c, t->cmd, 0, 0);
}
/* a column of the result */
typedef struct rcolx { adef def; int src; int fi; achap *chap; } rcolx;

static ars *shape_eval(acn *c, shp *s, const char *text)
{
    ars *par = 0;                               /* the rows of the parent (APPEND), the child (COMPUTE) */
    const char *alias = s->alias ? s->alias : s->tab.sub ? s->tab.sub->alias : 0;   /* SHAPE (SHAPE ... AS x ...) COMPUTE x */
    if (s->tab.cmd || s->tab.sub) { par = shape_tab(c, &s->tab, text); if (!par) return 0; }
    int nf = s->nf;
    ars **kids = calloc((size_t)(nf ? nf : 1), sizeof(ars *));
    achap **chs = calloc((size_t)(nf ? nf : 1), sizeof(achap *));
    char (*chapnm)[16] = calloc((size_t)(nf ? nf : 1), 16);   /* "Chapter<n>": the names live until the end */
    rcolx *cols = 0; int ncols = 0, cap = 0;
    ars *res = 0;
    int ok = 1;
#define ADDCOL(d_, src_, fi_) do { if (ncols == cap) { cap = cap ? cap * 2 : 16; cols = realloc(cols, sizeof(rcolx) * (size_t)cap); } \
        cols[ncols].def = (d_); cols[ncols].src = (src_); cols[ncols].fi = (fi_); cols[ncols].chap = 0; ncols++; } while (0)
    /* the groups of COMPUTE: one row for each different value of the BY fields */
    int ng = 0; int *grp = 0; int *bycol = 0;    /* grp[g]: a row of the child that has the group's values */
    drs *pb = par ? BASE(par) : 0;
    if (s->compute) {
        if (!par) { raise(-2147217900, SRC_SHAPE, "Data shape command text contains a syntax error."); ok = 0; goto done; }
        bycol = calloc((size_t)(s->nby ? s->nby : 1), sizeof(int));
        for (int k = 0; k < s->nby; k++) {
            bycol[k] = col_named(par, s->by[k]);
            if (bycol[k] < 0) {
                char msg[400]; snprintf(msg, sizeof msg, "Column (%s) does not exist in the appropriate rowset.", s->by[k]);
                raise(-2147217900, SRC_SHAPE, msg); ok = 0; goto done;
            }
        }
        grp = calloc((size_t)(pb->d->n + 1), sizeof(int));
        for (int i = 0; i < pb->d->n; i++) {
            if (pb->d->r[i].del) continue;
            int g;
            for (g = 0; g < ng; g++) {
                int same = 1;
                for (int k = 0; k < s->nby && same; k++) same = veq(pb->d->r[grp[g]].v[bycol[k]], pb->d->r[i].v[bycol[k]]);
                if (same) break;
            }
            if (g == ng) grp[ng++] = i;
            if (!s->nby) break;
        }
        if (!s->nby && !ng) ng = 1, grp[0] = -1;
        /* in the order of the BY values */
        for (int a = 1; a < ng; a++)
            for (int b = a; b > 0; b--) {
                int lt = 0;
                for (int k = 0; k < s->nby; k++) {
                    vbvar x = pb->d->r[grp[b]].v[bycol[k]], y = pb->d->r[grp[b - 1]].v[bycol[k]];
                    if (vless(x, y)) { lt = 1; break; }
                    if (vless(y, x)) break;
                }
                if (!lt) break;
                int t = grp[b]; grp[b] = grp[b - 1]; grp[b - 1] = t;
            }
    } else if (par) {
        for (int i = 0; i < par->ndefs; i++) ADDCOL(par->defs[i], 1, i);
    }
    /* the fields of APPEND / COMPUTE */
    for (int i = 0; i < nf; i++) {
        sfld *f = &s->f[i];
        adef d; memset(&d, 0, sizeof d);
        if (f->kind == SF_REL) {
            if (!par) { raise(-2147217900, SRC_SHAPE, "Data shape command text contains a syntax error."); ok = 0; goto done; }
            achap *a = calloc(1, sizeof *a); a->refs = 1;
            chs[i] = a;
            int haspar = 0;
            for (int k = 0; k < f->n; k++) if (f->par[k]) haspar = 1;
            for (int k = 0; k < f->n; k++) {
                a->pcol[k] = col_named(par, f->pn[k]);
                if (a->pcol[k] < 0) {
                    char msg[400]; snprintf(msg, sizeof msg, "Column (%s) does not exist in the appropriate rowset.", f->pn[k]);
                    raise(-2147217900, SRC_SHAPE, msg); ok = 0; goto done;
                }
                a->pidx[k] = f->par[k] - 1;
            }
            a->n = f->n;
            if (haspar) {
                if (!f->tab.cmd) { raise(-2147217900, SRC_SHAPE, "Data shape command text contains a syntax error."); ok = 0; goto done; }
                a->psql = strdup(f->tab.cmd);
            } else {
                ars *k = shape_tab(c, &f->tab, text);
                if (!k) { ok = 0; goto done; }
                kids[i] = k;
                a->child = (vbobj)k; vb_oaddref((vbobj)k);
                for (int q = 0; q < f->n; q++) {
                    a->ccol[q] = col_named(k, f->cn[q]);
                    if (a->ccol[q] < 0) {
                        char msg[400]; snprintf(msg, sizeof msg, "Column (%s) does not exist in the appropriate rowset.", f->cn[q]);
                        raise(-2147217900, SRC_SHAPE, msg); ok = 0; goto done;
                    }
                }
            }
            snprintf(chapnm[i], 16, "Chapter%d", i + 1);
            d.name = f->alias ? f->alias : f->talias ? f->talias : chapnm[i];
            if (f->alias) d.name = f->alias;
            d.type = 136; d.size = 4; d.attrs = 0x2000 | adFldFixed; d.prec = 255; d.scale = 255;
            ADDCOL(d, 2, i); cols[ncols - 1].chap = a;
        } else if (f->kind == SF_FIELD && s->compute) {
            const char *n = f->fn;
            if (alias && !strcasecmp(n, alias)) {     /* the child rows of the group */
                achap *a = calloc(1, sizeof *a); a->refs = 1;
                a->child = (vbobj)par; vb_oaddref((vbobj)par);
                a->n = s->nby;
                for (int k = 0; k < s->nby; k++) { a->ccol[k] = bycol[k]; a->pcol[k] = -1 - k; }  /* fixed below */
                chs[i] = a;
                d.name = f->alias ? f->alias : (char *)alias;
                d.type = 136; d.size = 4; d.attrs = 0x2000 | adFldFixed; d.prec = 255; d.scale = 255;
                ADDCOL(d, 2, i); cols[ncols - 1].chap = a;
            } else {
                char msg[400]; snprintf(msg, sizeof msg, "Column (%s) does not exist in the appropriate rowset.", n);
                raise(-2147217900, SRC_SHAPE, msg); ok = 0; goto done;
            }
        } else if (f->kind == SF_AGG) {
            d.name = f->alias ? f->alias : f->fn;
            d.type = 0; d.prec = 255; d.scale = 255;
            ADDCOL(d, 3, i);
        } else if (f->kind == SF_CALC) {
            d.name = f->alias ? f->alias : "Calc";
            ADDCOL(d, 4, i);
        } else if (f->kind == SF_NEW) {
            d.name = f->alias ? f->alias : "New";
            d.type = f->newtype; d.size = f->newsize; d.attrs = 0x40 | 0x20 | 0x4 | (f->newtype == 202 || f->newtype == 200 || f->newtype == 130 ? 0 : 0x10);
            d.prec = 255; d.scale = 255;
            ADDCOL(d, 5, i);
        } else { raise(-2147217900, SRC_SHAPE, "Data shape command text contains a syntax error."); ok = 0; goto done; }
    }
    if (s->compute)
        for (int k = 0; k < s->nby; k++) {      /* the BY fields, after the others */
            adef d = par->defs[bycol[k]];
            d.attrs &= ~adFldUnknownUpdatable;
            ADDCOL(d, 6, k);
        }
    /* the chapters of COMPUTE: parent column = the BY field of the result */
    for (int i = 0; i < nf; i++)
        if (chs[i] && s->compute)
            for (int k = 0; k < chs[i]->n; k++)
                for (int q = 0; q < ncols; q++) if (cols[q].src == 6 && cols[q].fi == k) chs[i]->pcol[k] = q;
    /* the aggregates: their chapter, field and type */
    int nrows = s->compute ? ng : 0;
    if (!s->compute && pb) for (int i = 0; i < pb->d->n; i++) if (!pb->d->r[i].del) nrows++;
    int *aggch = calloc((size_t)(ncols ? ncols : 1), sizeof(int)), *aggf = calloc((size_t)(ncols ? ncols : 1), sizeof(int));
    for (int q = 0; q < ncols; q++) {
        if (cols[q].src != 3) continue;
        sfld *f = &s->f[cols[q].fi];
        int ci = -1;
        for (int x = 0; x < ncols; x++) if (cols[x].chap && !strcasecmp(cols[x].def.name, f->qal)) ci = x;
        if (ci < 0 && s->compute && alias && !strcasecmp(f->qal, alias))
            for (int x = 0; x < ncols; x++) if (cols[x].chap) ci = x;
        if (ci < 0 || !cols[ci].chap->child) {
            char msg[400]; snprintf(msg, sizeof msg, "Column (%s) does not exist in the appropriate rowset.", f->qal);
            raise(-2147217900, SRC_SHAPE, msg); ok = 0; free(aggch); free(aggf); goto done;
        }
        ars *k = ARS(cols[ci].chap->child);
        int fc = f->qfld ? col_named(k, f->qfld) : -1;
        if (f->qfld && fc < 0) {
            char msg[400]; snprintf(msg, sizeof msg, "Column (%s) does not exist in the appropriate rowset.", f->qfld);
            raise(-2147217900, SRC_SHAPE, msg); ok = 0; free(aggch); free(aggf); goto done;
        }
        aggch[q] = ci; aggf[q] = fc;
        adef *d = &cols[q].def;
        const char *fn = f->fn;
        int ft = fc >= 0 ? k->defs[fc].type : adInteger;
        if (!strcmp(fn, "COUNT")) { d->type = adInteger; d->size = 4; d->attrs = 0x40 | 0x10; }
        else if (!strcmp(fn, "SUM") || !strcmp(fn, "AVG")) { d->type = ft == adCurrency ? adCurrency : adDouble; d->size = 8; d->attrs = 0x40 | 0x10; }
        else if (!strcmp(fn, "STDEV")) { d->type = adDouble; d->size = 8; d->attrs = 0x40 | 0x10; }
        else { *d = k->defs[fc]; d->name = f->alias ? f->alias : (char *)fn; d->attrs &= ~adFldUnknownUpdatable; }
        d->prec = precision_of(d->type); d->scale = 255;
    }
    /* the values, row by row */
    vbvar **vals = calloc((size_t)(nrows ? nrows : 1), sizeof(vbvar *));
    int r0 = 0;
    for (int i = 0; s->compute ? r0 < ng : (pb && i < pb->d->n); i++) {
        if (!s->compute && pb->d->r[i].del) continue;
        vbvar *row = calloc((size_t)(ncols ? ncols : 1), sizeof(vbvar));
        for (int q = 0; q < ncols; q++) {
            row[q] = vb_v_null();
            if (cols[q].src == 1) vb_vset(&row[q], pb->d->r[i].v[cols[q].fi]);
            else if (cols[q].src == 6) { if (grp[r0] >= 0) vb_vset(&row[q], pb->d->r[grp[r0]].v[bycol[cols[q].fi]]); }
            else if (cols[q].src == 2) row[q] = vb_v_l(r0 + 1);
        }
        for (int q = 0; q < ncols; q++) {
            if (cols[q].src != 3) continue;
            achap *a = cols[aggch[q]].chap;
            vbvar pv[16];
            for (int k = 0; k < a->n; k++) pv[k] = row[a->pcol[k]];
            vb_vset(&row[q], aggregate(s->f[cols[q].fi].fn, ARS(a->child), aggf[q], a, pv));
        }
        vals[r0++] = row;
        if (s->compute) i = -1 + 0;             /* COMPUTE: by group, not by row */
    }
    nrows = r0;
    free(aggch); free(aggf);
    /* CALC: the expressions, and their types from their values */

    {
        ars tmp; memset(&tmp, 0, sizeof tmp);
        adef *td = calloc((size_t)(ncols ? ncols : 1), sizeof(adef));
        for (int q = 0; q < ncols; q++) td[q] = cols[q].def;
        tmp.defs = td; tmp.ndefs = ncols;
        for (int q = 0; q < ncols; q++) {
            if (cols[q].src != 4) continue;
            int vt = VT_NULL;
            for (int k = 0; k < nrows; k++) {
                int err = 0;
                vbvar x = calc_eval(s->f[cols[q].fi].fn, &tmp, vals[k], ncols, &err);
                if (err) { free(td); for (int z = 0; z < nrows; z++) { for (int y = 0; y < ncols; y++) vb_vfree(&vals[z][y]); free(vals[z]); } free(vals); ok = 0; goto done; }
                vb_vset(&vals[k][q], x);
                if (vt == VT_NULL && x.vt != VT_NULL && x.vt != VT_EMPTY) vt = x.vt;
            }
            adef *d = &cols[q].def;
            switch (vt) {
            case VT_I2: d->type = adSmallInt; d->size = 2; break;
            case VT_I4: d->type = adInteger; d->size = 4; break;
            case VT_R4: d->type = adSingle; d->size = 4; break;
            case VT_R8: d->type = adDouble; d->size = 8; break;
            case VT_CY: d->type = adCurrency; d->size = 8; break;
            case VT_DATE: d->type = adDate; d->size = 8; break;
            case VT_BOOL: d->type = adBoolean; d->size = 2; break;
            case VT_UI1: d->type = adUnsignedTinyInt; d->size = 1; break;
            case VT_BSTR: d->type = adVarWChar; d->size = 254; break;
            default: d->type = adVariant; d->size = 16; break;
            }
            d->attrs = d->type == adVarWChar ? 0x40 : 0x40 | 0x10;
            d->prec = precision_of(d->type); d->scale = 255;
        }
        free(td);
    }
    /* the recordset */
    res = (ars *)rs_new_for(c);
    {
        adef *defs = calloc((size_t)(ncols ? ncols : 1), sizeof(adef));
        for (int q = 0; q < ncols; q++) defs[q] = cols[q].def;
        int e = make_memory(res, defs, ncols);
        free(defs);
        if (e) { for (int z = 0; z < nrows; z++) { for (int y = 0; y < ncols; y++) vb_vfree(&vals[z][y]); free(vals[z]); } free(vals); vb_orelease((vbobj)res); res = 0; ok = 0; goto done; }
        for (int q = 0; q < ncols; q++) { res->defs[q].prec = cols[q].def.prec; res->defs[q].scale = cols[q].def.scale; }
        drs *m = BASE(res);
        jbuf sq = { 0 };
        jb_s(&sq, "INSERT INTO lbrs VALUES (");
        for (int q = 0; q < ncols; q++) jb_f(&sq, q ? ", ?%d" : "?%d", q + 1);
        jb_s(&sq, ")");
        sqlite3_stmt *ins = 0;
        sqlite3_prepare_v2(DB(res->db)->db, sq.p, -1, &ins, 0);
        jb_free(&sq);
        for (int z = 0; z < nrows && ins; z++) {
            sqlite3_reset(ins);
            for (int q = 0; q < ncols; q++) {
                if (cols[q].chap) sqlite3_bind_null(ins, q + 1);
                else jet_bind(ins, q + 1, vals[z][q], m->cols[q].type);
            }
            sqlite3_step(ins);
        }
        if (ins) sqlite3_finalize(ins);
        for (int z = 0; z < nrows; z++) { for (int y = 0; y < ncols; y++) vb_vfree(&vals[z][y]); free(vals[z]); }
        free(vals);
        int f2;
        rs_member(res->base, "requery", VBM_METHOD, 0, 0, &f2);
        if (vb_errpend) vb_err_clear_pending();
        rebuild_view(res);
        res->exact = 1;
        res->cursortype = adOpenStatic; res->cursorloc = adUseClient;
        res->chaps = calloc((size_t)(ncols ? ncols : 1), sizeof(achap *));
        for (int q = 0; q < ncols; q++) if (cols[q].chap) { res->chaps[q] = cols[q].chap; cols[q].chap->refs++; }
        /* changes of the parent's own columns go to the parent's table */
        if (par && !s->compute && par->ot) {
            otarget *t = calloc(1, sizeof *t), *pt = par->ot;
            t->table = strdup(pt->table); t->ncols = ncols; t->autocol = pt->autocol;
            t->cols = calloc((size_t)(ncols ? ncols : 1), sizeof(char *));
            t->iskey = calloc((size_t)(ncols ? ncols : 1), sizeof(int));
            for (int q = 0; q < ncols; q++)
                if (cols[q].src == 1 && cols[q].fi < pt->ncols) {
                    t->cols[q] = pt->cols[cols[q].fi] ? strdup(pt->cols[cols[q].fi]) : 0;
                    t->iskey[q] = pt->iskey[cols[q].fi];
                }
            res->ot = t;
            DB(res->db)->remote = cursor_write; DB(res->db)->remote_ctx = res;
        }
        res->state = 1;
    }
done:
    for (int i = 0; i < nf; i++) { if (kids[i]) vb_orelease((vbobj)kids[i]); chap_free(chs[i]); }
    free(kids); free(chs); free(cols); free(grp); free(bycol); free(chapnm);
    if (par) vb_orelease((vbobj)par);
    (void)ok;
#undef ADDCOL
    return res;
}

/* Recordset.Open / Connection.Execute of a SHAPE command */
static int shape_open(ars *r, acn *c, const char *text, int locktype)
{
    spar sp = { text, text, 0, 0 };
    shp *s = sp_shape(&sp);
    sp_ws(&sp);
    if (s && *sp.p) { shp_free(s); s = 0; if (!sp.err) { sp.err = 1; sp.errpos = (int)(sp.p - text); } }
    free(r->source); r->source = strdup(text);
    if (!s) { shape_syntax_error(text, sp.errpos); return -1; }
    alist_clear(c->errs);
    ars *res = shape_eval(c, s, text);
    shp_free(s);
    if (!res) return -1;
    /* r takes over the rows of res */
    rs_drop_rows(r);
    r->db = res->db; res->db = 0;
    r->base = res->base; res->base = 0;
    r->view = res->view; res->view = 0;
    r->memdb = 1; res->memdb = 0;
    for (int i = 0; i < r->ndefs; i++) free(r->defs[i].name);
    free(r->defs);
    r->defs = res->defs; r->ndefs = res->ndefs; res->defs = 0; res->ndefs = 0;
    r->chaps = res->chaps; res->chaps = 0;
    r->ot = res->ot; res->ot = 0;
    r->exact = 1;
    if (DB(r->db)->remote_ctx == res) DB(r->db)->remote_ctx = r;
    vb_orelease((vbobj)res);
    r->cursortype = adOpenStatic; r->cursorloc = adUseClient;
    r->locktype = locktype;
    r->state = 1;
    free(r->source); r->source = strdup(text);
    go_first(VW(r));
    return 0;
}

/* ---------------------------------------------------------------- chapters */
static void vlit(jbuf *b, vbvar v)              /* a value in the text of a Filter */
{
    char t[64];
    switch (v.vt) {
    case VT_I2: case VT_I4: case VT_UI1: jb_f(b, "%d", (int)vb_v2l(v)); break;
    case VT_R4: case VT_R8: case VT_CY: snprintf(t, sizeof t, "%.17g", vb_v2d(v)); jb_s(b, t); break;
    case VT_BOOL: jb_s(b, v.u.b ? "-1" : "0"); break;
    case VT_DATE: {
        vb_date d = vb_v2date(v);
        jb_f(b, "#%04d-%02d-%02d %02d:%02d:%02d#", vb_year(d), vb_month(d), vb_day(d), vb_hour(d), vb_minute(d), vb_second(d));
        break;
    }
    default: {
        const char *s = vb_stoutf8(vb_v2s(v));
        jb_s(b, "'");
        for (; *s; s++) { if (*s == '\'') jb_s(b, "'"); jb_put(b, s, 1); }
        jb_s(b, "'");
    }
    }
}
/* the rows of the chapter for the current row of its parent */
static int chap_refilter(ars *ch)
{
    ars *p = ARS(ch->chparent);
    drs *pv = VW(p);
    achap *a = p->chaps[ch->chcol];
    ch->chrow = pv->cur;
    vbvar *row = alive(pv, pv->cur) ? pv->d->r[pv->cur].v : 0;
    if (a->psql) {                              /* RELATE ... TO PARAMETER: the command again */
        acn *c = cn_of(p);
        vbvar vals[16]; int nv = 0;
        for (int k = 0; k < a->n; k++) { int ix = a->pidx[k]; if (ix >= 0 && ix < 16) { vals[ix] = row ? row[a->pcol[k]] : vb_v_null(); if (ix + 1 > nv) nv = ix + 1; } }
        ars *k = c ? shape_exec(c, a->psql, vals, nv) : 0;
        if (!k) return -1;
        int ct = ch->cursortype, lt = ch->locktype;
        rs_drop_rows(ch);
        ch->db = k->db; k->db = 0; ch->base = k->base; k->base = 0; ch->view = k->view; k->view = 0;
        ch->memdb = 1; k->memdb = 0;
        for (int i = 0; i < ch->ndefs; i++) free(ch->defs[i].name);
        free(ch->defs);
        ch->defs = k->defs; ch->ndefs = k->ndefs; k->defs = 0; k->ndefs = 0;
        ch->ot = k->ot; k->ot = 0;
        if (DB(ch->db)->remote_ctx == k) DB(ch->db)->remote_ctx = ch;
        vb_orelease((vbobj)k);
        ch->cursortype = ct; ch->locktype = lt;
        ch->chrow = pv->cur;
        go_first(VW(ch));
        return 0;
    }
    jbuf b = { 0 };
    ars *m = ARS(a->child);
    for (int k = 0; k < a->n; k++) {
        if (b.n) jb_s(&b, " AND ");
        vbvar x = row ? row[a->pcol[k]] : vb_v_null();
        jb_s(&b, "["); jb_s(&b, BASE(m)->cols[a->ccol[k]].name); jb_s(&b, "]");
        if (!row) { jb_s(&b, " IS NULL AND 1 = 0"); continue; }
        if (x.vt == VT_NULL || x.vt == VT_EMPTY) jb_s(&b, " IS NULL");
        else { jb_s(&b, " = "); vlit(&b, x); }
    }
    /* the rows of the chapter: a recordset of their own over the rows of the master (not all of
       them, filtered: a hierarchy of thousands of rows asks for a chapter at every row) */
    if (!a->indexed && a->n) {
        jbuf ix = { 0 };
        jb_f(&ix, "CREATE INDEX IF NOT EXISTS lbch%p ON lbrs (", (void *)a);
        for (int k = 0; k < a->n; k++) { if (k) jb_s(&ix, ", "); jb_s(&ix, "\""); jb_s(&ix, BASE(m)->cols[a->ccol[k]].name); jb_s(&ix, "\""); }
        jb_s(&ix, ")");
        sqlite3_exec(DB(m->db)->db, ix.p, 0, 0, 0);
        jb_free(&ix);
        a->indexed = 1;
    }
    jbuf q = { 0 };
    jb_s(&q, "SELECT * FROM lbrs");
    if (b.n) { jb_s(&q, " WHERE "); jb_s(&q, b.p); }
    jb_free(&b);
    vbobj nb = rs_open(DB(ch->db), q.p, dbOpenDynaset, 0, 0, 0, 0, 0);
    jb_free(&q);
    if (!nb) { if (vb_errpend) vb_err_clear_pending(); return -1; }
    vb_oaddref(nb);
    if (ch->view && ch->view != ch->base) { rs_close((drs *)vb_oid(ch->view)); vb_orelease(ch->view); }
    else if (ch->view) vb_orelease(ch->view);
    if (ch->base) { rs_close((drs *)vb_oid(ch->base)); vb_orelease(ch->base); }
    ch->base = nb; ch->view = 0;
    free(ch->chcrit); ch->chcrit = 0;
    rebuild_view(ch);
    return 0;
}
static vbvar chap_value(ars *p, int col)
{
    achap *a = p->chaps[col];
    ars *ch = (ars *)rs_new_for(cn_of(p));
    ch->chparent = (vbobj)p; vb_oaddref((vbobj)p);
    ch->chcol = col;
    ch->cursortype = adOpenStatic; ch->cursorloc = adUseClient; ch->locktype = p->locktype;
    ch->exact = 1;
    if (!a->psql) {
        ars *m = ARS(a->child);
        ch->db = m->db; vb_oaddref(ch->db);
        ch->memdb = 1; ch->shared = 1;
        ch->base = 0;                           /* chap_refilter opens it */
        ch->defs = malloc(sizeof(adef) * (size_t)(m->ndefs ? m->ndefs : 1));
        for (int i = 0; i < m->ndefs; i++) { ch->defs[i] = m->defs[i]; ch->defs[i].name = strdup(m->defs[i].name); }
        ch->ndefs = m->ndefs;
        if (m->chaps) {
            ch->chaps = calloc((size_t)(m->ndefs ? m->ndefs : 1), sizeof(achap *));
            for (int i = 0; i < m->ndefs; i++) if (m->chaps[i]) { ch->chaps[i] = m->chaps[i]; m->chaps[i]->refs++; }
        }
    }
    ch->state = 1;
    if (chap_refilter(ch)) { vb_orelease((vbobj)ch); return vb_v_empty(); }
    vbvar v = retobj((vbobj)ch);
    vb_orelease((vbobj)ch);
    return v;
}
/* StayInSync: a chapter follows its parent */
static void chap_sync(ars *ch)
{
    if (!ch->chparent || !ch->state) return;
    ars *p = ARS(ch->chparent);
    if (!p->state || !p->view) return;
    if (VW(p)->cur != ch->chrow) chap_refilter(ch);
}

/* ---------------------------------------------------------------- CALC: an expression of VBA */
typedef struct cx { const char *p; ars *r; vbvar *row; int n; int err; } cx;
static vbvar cx_expr(cx *x);
static void cx_ws(cx *x) { while (*x->p == ' ' || *x->p == '\t') x->p++; }
static int cx_kw(cx *x, const char *w)
{
    cx_ws(x);
    size_t n = strlen(w);
    if (strncasecmp(x->p, w, n) || isalnum((unsigned char)x->p[n]) || x->p[n] == '_') return 0;
    x->p += n;
    return 1;
}
static vbvar cx_fail(cx *x) { if (!x->err) { x->err = 1; raise(-2147217900, SRC_SHAPE, "Data shape command text contains a syntax error."); } return vb_v_null(); }
static vbvar cx_call(cx *x, const char *fn, vbvar *a, int n)
{
#define NA(k) if (n != (k)) return cx_fail(x)
#define ISN(v) ((v).vt == VT_NULL)
    if (!strcasecmp(fn, "IIf")) { NA(3); return vb_v_truth(a[0]) ? a[1] : a[2]; }
    if (!strcasecmp(fn, "IsNull")) { NA(1); return vb_v_bool(ISN(a[0]) ? -1 : 0); }
    if (!strcasecmp(fn, "Len")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_l(vb_v2s(a[0]).len); }
    if (!strcasecmp(fn, "Left") || !strcasecmp(fn, "Right")) { NA(2); if (ISN(a[0])) return a[0]; vbs s = vb_v2s(a[0]); return vb_v_s(fn[0] == 'L' || fn[0] == 'l' ? vb_left(s, vb_v2l(a[1])) : vb_right(s, vb_v2l(a[1]))); }
    if (!strcasecmp(fn, "Mid")) { if (n < 2 || n > 3) return cx_fail(x); if (ISN(a[0])) return a[0]; return vb_v_s(vb_mid(vb_v2s(a[0]), vb_v2l(a[1]), n == 3 ? vb_v2l(a[2]) : -1)); }
    if (!strcasecmp(fn, "UCase")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_s(vb_ucase(vb_v2s(a[0]))); }
    if (!strcasecmp(fn, "LCase")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_s(vb_lcase(vb_v2s(a[0]))); }
    if (!strcasecmp(fn, "Trim")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_s(vb_trim(vb_v2s(a[0]))); }
    if (!strcasecmp(fn, "LTrim")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_s(vb_ltrim(vb_v2s(a[0]))); }
    if (!strcasecmp(fn, "RTrim")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_s(vb_rtrim(vb_v2s(a[0]))); }
    if (!strcasecmp(fn, "InStr")) { if (n == 2) return vb_v_l(vb_instr(1, vb_v2s(a[0]), vb_v2s(a[1]), 0)); NA(3); return vb_v_l(vb_instr(vb_v2l(a[0]), vb_v2s(a[1]), vb_v2s(a[2]), 0)); }
    if (!strcasecmp(fn, "Space")) { NA(1); return vb_v_s(vb_space(vb_v2l(a[0]))); }
    if (!strcasecmp(fn, "CStr")) { NA(1); return vb_v_s(vb_v2s(a[0])); }
    if (!strcasecmp(fn, "CLng")) { NA(1); return vb_v_l(vb_v2l(a[0])); }
    if (!strcasecmp(fn, "CInt")) { NA(1); return vb_v_i((vb_i16)vb_v2l(a[0])); }
    if (!strcasecmp(fn, "CDbl")) { NA(1); return vb_v_d(vb_v2d(a[0])); }
    if (!strcasecmp(fn, "CSng")) { NA(1); return vb_v_f(vb_v2f(a[0])); }
    if (!strcasecmp(fn, "CCur")) { NA(1); return vb_v_c(vb_v2c(a[0])); }
    if (!strcasecmp(fn, "CDate")) { NA(1); return vb_v_date(vb_v2date(a[0])); }
    if (!strcasecmp(fn, "CBool")) { NA(1); return vb_v_bool(vb_v2bool(a[0])); }
    if (!strcasecmp(fn, "Val")) { NA(1); return vb_v_d(vb_val(vb_v2s(a[0]))); }
    if (!strcasecmp(fn, "Abs")) { NA(1); if (ISN(a[0])) return a[0]; vbvar z = vb_v_cmp(a[0], vb_v_l(0), 2, 0); return vb_v_truth(z) ? vb_v_neg(a[0]) : a[0]; }
    if (!strcasecmp(fn, "Int") || !strcasecmp(fn, "Fix")) { NA(1); if (ISN(a[0])) return a[0]; double d = vb_v2d(a[0]); d = (fn[0] == 'I' || fn[0] == 'i') ? floor(d) : (d < 0 ? ceil(d) : floor(d)); return vb_v_d(d); }
    if (!strcasecmp(fn, "Round")) { if (n < 1 || n > 2) return cx_fail(x); if (ISN(a[0])) return a[0]; return vb_v_d(vb_round(vb_v2d(a[0]), n == 2 ? vb_v2l(a[1]) : 0)); }
    if (!strcasecmp(fn, "Sqr")) { NA(1); return vb_v_d(sqrt(vb_v2d(a[0]))); }
    if (!strcasecmp(fn, "Year")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_i((vb_i16)vb_year(vb_v2date(a[0]))); }
    if (!strcasecmp(fn, "Month")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_i((vb_i16)vb_month(vb_v2date(a[0]))); }
    if (!strcasecmp(fn, "Day")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_i((vb_i16)vb_day(vb_v2date(a[0]))); }
    if (!strcasecmp(fn, "Hour")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_i((vb_i16)vb_hour(vb_v2date(a[0]))); }
    if (!strcasecmp(fn, "Minute")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_i((vb_i16)vb_minute(vb_v2date(a[0]))); }
    if (!strcasecmp(fn, "Second")) { NA(1); return ISN(a[0]) ? a[0] : vb_v_i((vb_i16)vb_second(vb_v2date(a[0]))); }
    if (!strcasecmp(fn, "Now")) { NA(0); return vb_v_date(vb_now()); }
    if (!strcasecmp(fn, "Date")) { NA(0); return vb_v_date(floor(vb_now())); }
    if (!strcasecmp(fn, "Format")) { if (n < 1 || n > 2) return cx_fail(x); if (ISN(a[0])) return a[0]; return vb_v_s(vb_format(a[0], n == 2 ? vb_v2s(a[1]) : vb_sfromutf8(""), n == 2, 1, 1)); }
#undef NA
#undef ISN
    return cx_fail(x);
}
static vbvar cx_atom(cx *x)
{
    cx_ws(x);
    const char *p = x->p;
    if (*p == '(') { x->p++; vbvar v = cx_expr(x); cx_ws(x); if (*x->p != ')') return cx_fail(x); x->p++; return v; }
    if (*p == '-') { x->p++; return vb_v_neg(cx_atom(x)); }
    if (*p == '+') { x->p++; return cx_atom(x); }
    if (*p == '"') {
        jbuf b = { 0 };
        p++;
        for (;;) {
            if (!*p) return cx_fail(x);
            if (*p == '"') { if (p[1] == '"') { jb_put(&b, "\"", 1); p += 2; continue; } p++; break; }
            jb_put(&b, p++, 1);
        }
        x->p = p;
        vbvar v = vb_v_s(vb_sfromutf8(b.p ? b.p : ""));
        jb_free(&b);
        return v;
    }
    if (*p == '#') {
        const char *e = strchr(p + 1, '#');
        if (!e) return cx_fail(x);
        char t[64]; snprintf(t, sizeof t, "%.*s", (int)(e - p - 1), p + 1);
        x->p = e + 1;
        return vb_v_date(vb_v2date(vb_v_s(vb_sfromutf8(t))));
    }
    if (isdigit((unsigned char)*p) || (*p == '.' && isdigit((unsigned char)p[1]))) {
        char *e; double d = strtod(p, &e);
        x->p = e;
        int integer = 1;
        for (const char *q = p; q < e; q++) if (*q == '.' || *q == 'e' || *q == 'E') integer = 0;
        if (integer && d >= -32768 && d <= 32767) return vb_v_i((vb_i16)d);
        if (integer && d >= -2147483648.0 && d <= 2147483647.0) return vb_v_l((vb_i32)d);
        return vb_v_d(d);
    }
    if (*p == '[' || isalpha((unsigned char)*p) || *p == '_') {
        char nm[256]; int k = 0;
        if (*p == '[') { p++; while (*p && *p != ']' && k < 255) nm[k++] = *p++; if (*p == ']') p++; }
        else while ((isalnum((unsigned char)*p) || *p == '_') && k < 255) nm[k++] = *p++;
        nm[k] = 0;
        x->p = p;
        if (!strcasecmp(nm, "True")) return vb_v_bool(-1);
        if (!strcasecmp(nm, "False")) return vb_v_bool(0);
        if (!strcasecmp(nm, "Null")) return vb_v_null();
        cx_ws(x);
        if (*x->p == '(') {
            x->p++;
            vbvar a[8]; int n = 0;
            cx_ws(x);
            if (*x->p != ')')
                do { if (n == 8) return cx_fail(x); a[n++] = cx_expr(x); if (x->err) return vb_v_null(); cx_ws(x); } while (*x->p == ',' && x->p++);
            if (*x->p != ')') return cx_fail(x);
            x->p++;
            return cx_call(x, nm, a, n);
        }
        if (!strcasecmp(nm, "Now") || !strcasecmp(nm, "Date")) return cx_call(x, nm, 0, 0);
        for (int i = 0; i < x->n; i++) if (!strcasecmp(x->r->defs[i].name, nm)) return x->row[i];
        if (!x->err) {
            x->err = 1;
            char msg[400]; snprintf(msg, sizeof msg, "Column (%s) does not exist in the appropriate rowset.", nm);
            raise(-2147217900, SRC_SHAPE, msg);
        }
        return vb_v_null();
    }
    return cx_fail(x);
}
static vbvar cx_pow(cx *x) { vbvar a = cx_atom(x); while (!x->err) { cx_ws(x); if (*x->p != '^') break; x->p++; a = vb_v_pow(a, cx_atom(x)); } return a; }
static vbvar cx_mul(cx *x)
{
    vbvar a = cx_pow(x);
    while (!x->err) {
        cx_ws(x);
        if (*x->p == '*') { x->p++; a = vb_v_mul(a, cx_pow(x)); }
        else if (*x->p == '/') { x->p++; a = vb_v_div(a, cx_pow(x)); }
        else if (*x->p == '\\') { x->p++; a = vb_v_idiv(a, cx_pow(x)); }
        else if (cx_kw(x, "Mod")) a = vb_v_mod(a, cx_pow(x));
        else break;
    }
    return a;
}
static vbvar cx_add(cx *x)
{
    vbvar a = cx_mul(x);
    while (!x->err) {
        cx_ws(x);
        if (*x->p == '+') { x->p++; a = vb_v_add(a, cx_mul(x)); }
        else if (*x->p == '-') { x->p++; a = vb_v_sub(a, cx_mul(x)); }
        else break;
    }
    return a;
}
static vbvar cx_cat(cx *x) { vbvar a = cx_add(x); while (!x->err) { cx_ws(x); if (*x->p != '&') break; x->p++; a = vb_v_cat(a, cx_add(x)); } return a; }
static vbvar cx_cmp(cx *x)
{
    vbvar a = cx_cat(x);
    while (!x->err) {
        cx_ws(x);
        int op = -1;
        if (!strncmp(x->p, "<>", 2)) { op = 1; x->p += 2; }
        else if (!strncmp(x->p, "<=", 2)) { op = 4; x->p += 2; }
        else if (!strncmp(x->p, ">=", 2)) { op = 5; x->p += 2; }
        else if (*x->p == '=') { op = 0; x->p++; }
        else if (*x->p == '<') { op = 2; x->p++; }
        else if (*x->p == '>') { op = 3; x->p++; }
        if (op < 0) break;
        a = vb_v_cmp(a, cx_cat(x), op, 0);
    }
    return a;
}
static vbvar cx_not(cx *x) { if (cx_kw(x, "Not")) return vb_v_not(cx_not(x)); return cx_cmp(x); }
static vbvar cx_and(cx *x) { vbvar a = cx_not(x); while (!x->err && cx_kw(x, "And")) a = vb_v_and(a, cx_not(x)); return a; }
static vbvar cx_expr(cx *x)
{
    vbvar a = cx_and(x);
    while (!x->err) {
        if (cx_kw(x, "Or")) a = vb_v_or(a, cx_and(x));
        else if (cx_kw(x, "Xor")) a = vb_v_xor(a, cx_and(x));
        else break;
    }
    return a;
}
static vbvar calc_eval(const char *expr, ars *r, vbvar *row, int ncols, int *err)
{
    cx x = { expr, r, row, ncols, 0 };
    vb_i32 save = vb_errpend;
    vbvar v = cx_expr(&x);
    cx_ws(&x);
    if (!x.err && *x.p) cx_fail(&x);
    if (!x.err && vb_errpend && !save) {        /* an error of VBA (division by zero ...): Null */
        vb_err_clear_pending();
        v = vb_v_null();
    }
    *err = x.err;
    return x.err ? vb_v_null() : vb_vtmp(v);
}
