/* jetnew.c - a new, empty Jet 3 / Jet 4 database as a model (what CreateDatabase makes): page 0 and the four tables
 * of the catalog (MSysObjects, MSysACEs, MSysQueries, MSysRelationships) with their rows. Everything here is the
 * format as measured on the databases Jet makes (DEV-14: empty3, empty4, jetmany); nothing is copied from a file.
 * Jet 4 keeps the SIDs (MSysObjects.Owner, MSysACEs.SID) xored with a key of 2 bytes that follows from the creation
 * date in page 0 (how is not known): a new file takes a date whose key was measured. The Python prototype:
 * tools/jet/jetnew.py.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "jetfile.h"
#include <stdlib.h>
#include <string.h>
#include <time.h>

static const unsigned char JET4_DATE[8] = { 0xaa, 0x17, 0xdf, 0xb9, 0x41, 0xa2, 0xe6, 0x40 };   /* 2026-11-28 01:17:40 */
static const unsigned char JET4_KEY[2] = { 0xed, 0xa0 };
/* the other sort orders: a creation date and its key, as Jet made them (DEV-14/jetloc c_<LCID>.mdb); the order's code
 * page (page 0, 0x3C) */
static const struct { int lcid, cp; unsigned char date[8], key[2]; } JET4_ORDERS[] = {
    { 0x401, 1256, { 0x2a, 0xf5, 0x2b, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xe6, 0x74 } },
    { 0x402, 1251, { 0x0a, 0xa7, 0xb7, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x1c, 0x29 } },
    { 0x404, 950, { 0x94, 0x29, 0x48, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xa6, 0x47 } },
    { 0x405, 1250, { 0x10, 0x44, 0x79, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x3b, 0x60 } },
    { 0x406, 1252, { 0x19, 0xb3, 0x4d, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x4d, 0x4d } },
    { 0x407, 1252, { 0xaa, 0xc5, 0x1c, 0x33, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x43, 0x1c } },
    { 0x408, 1253, { 0xff, 0x65, 0xa9, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xce, 0x04 } },
    { 0x40a, 1252, { 0xb9, 0x8f, 0xa1, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x7e, 0x6f } },
    { 0x40b, 1252, { 0xcd, 0x3d, 0xe9, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xf3, 0xe5 } },
    { 0x40d, 1255, { 0x68, 0x6c, 0xc2, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xad, 0xf7 } },
    { 0x40e, 1250, { 0xe5, 0x3a, 0xdb, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xc3, 0xae } },
    { 0x40f, 1252, { 0x95, 0xbd, 0xf4, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x41, 0xe6 } },
    { 0x411, 932, { 0x19, 0xab, 0x08, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x3b, 0xc2 } },
    { 0x412, 949, { 0x3d, 0x63, 0x24, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x73, 0x86 } },
    { 0x413, 1252, { 0x4d, 0xbf, 0x93, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x78, 0xee } },
    { 0x414, 1252, { 0x36, 0x44, 0x02, 0x33, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xb9, 0xdc } },
    { 0x415, 1250, { 0x4f, 0x05, 0x66, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xdc, 0x89 } },
    { 0x418, 1250, { 0x36, 0x9e, 0x09, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x62, 0xdb } },
    { 0x419, 1251, { 0xb7, 0xb3, 0x60, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x3c, 0xf9 } },
    { 0x41a, 1250, { 0x96, 0x91, 0xf0, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xcc, 0x5f } },
    { 0x41b, 1250, { 0x59, 0x28, 0x22, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xa5, 0x52 } },
    { 0x41d, 1252, { 0x7a, 0xde, 0x3e, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x7d, 0xba } },
    { 0x41e, 874, { 0x96, 0xcb, 0xbc, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x21, 0xed } },
    { 0x41f, 1254, { 0xba, 0x83, 0xd8, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x6a, 0x76 } },
    { 0x422, 1251, { 0x7c, 0xb2, 0x3a, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x3a, 0xf7 } },
    { 0x424, 1250, { 0x3d, 0xb7, 0x81, 0x31, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xd4, 0x07 } },
    { 0x425, 1257, { 0xd5, 0x42, 0x53, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xd1, 0x41 } },
    { 0x426, 1257, { 0x1b, 0x0b, 0x6c, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x16, 0x9c } },
    { 0x427, 1257, { 0xde, 0x55, 0x85, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0x1f, 0x8e } },
    { 0x42a, 1258, { 0xb1, 0x16, 0x9f, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xaf, 0x66 } },
    { 0x804, 936, { 0x18, 0x93, 0x39, 0x30, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xec, 0x38 } },
    { 0xc0a, 1252, { 0xe7, 0xb4, 0xcf, 0x32, 0x4d, 0xa2, 0xe6, 0x40 }, { 0xb5, 0xff } },
};
static const unsigned char *cur_date = JET4_DATE, *cur_key = JET4_KEY;
enum { SID_SYSTEM = 0x0302, SID_ADMIN = 0x0103, SID_CREATOR = 0x0402, SID_ADMINS = 0x0102 };
#define ID_TABLES 0x0F000001u
#define ID_DATABASES 0x0F000002u
#define ID_RELATIONSHIPS 0x0F000003u
#define ID_MSYSDB 0x10000000u

static void P16(unsigned char *b, unsigned v) { b[0] = (unsigned char)v; b[1] = (unsigned char)(v >> 8); }
static void P32(unsigned char *b, uint32_t v) { b[0] = (unsigned char)v; b[1] = (unsigned char)(v >> 8); b[2] = (unsigned char)(v >> 16); b[3] = (unsigned char)(v >> 24); }

void jet_new_datekey(unsigned char date[8], unsigned char key[2]) { memcpy(date, cur_date, 8); memcpy(key, cur_key, 2); }

void jet_sid(int jet4, unsigned v, unsigned char out[2])
{
    out[0] = (unsigned char)v; out[1] = (unsigned char)(v >> 8);
    if (jet4) { out[0] ^= cur_key[0]; out[1] ^= cur_key[1]; }
}

static unsigned char *page0(int jet4)
{
    int ps = jet4 ? 4096 : 2048;
    unsigned char *p = calloc(1, (size_t)ps);
    p[1] = 1;
    memcpy(p + 4, "Standard Jet DB", 15);
    p[0x14] = jet4 ? 1 : 0;
    p[0x19] = 1;
    p[0x1C] = 1; p[0x1D] = 1;
    p[0x20] = 2; p[0x24] = 3; p[0x28] = 4; p[0x2C] = 5;
    int tab;
    if (jet4) {
        P16(p + 0x3C, 1252);
        double d;
        memcpy(&d, cur_date, 8);
        uint32_t mask = (uint32_t)(int)d;                   /* the empty password, masked */
        for (int i = 0; i < 10; i++) P32(p + 0x42 + 4 * i, mask);
        P32(p + 0x6A, 9801);
        P32(p + 0x6E, 0x409);
        memcpy(p + 0x72, cur_date, 8);
        P32(p + 0x98, 0x654);
        memcpy(p + 0x9C, "4.0", 3);
        tab = 0xE00;
    } else {
        P16(p + 0x3A, 0x409);
        P16(p + 0x3C, 1252);
        P16(p + 0x56, 9801);
        tab = 0x600;
    }
    P16(p + tab, 0x0101);
    for (int i = 1; i < 256; i++) P16(p + tab + 2 * i, 0x0100);
    return p;
}

typedef struct ccol { const char *name; int type, num, len, fixed; } ccol;
static const ccol C_OBJ[] = { { "Id", JT_LONG, 0, 4, 1 }, { "ParentId", JT_LONG, 1, 4, 1 }, { "Name", JT_TEXT, 2, 255, 0 },
    { "Type", JT_INT, 3, 2, 1 }, { "DateCreate", JT_DATETIME, 4, 8, 1 }, { "DateUpdate", JT_DATETIME, 5, 8, 1 },
    { "Owner", JT_BINARY, 6, 255, 0 }, { "Flags", JT_LONG, 7, 4, 1 }, { "Database", JT_MEMO, 8, 0, 0 },
    { "Connect", JT_MEMO, 9, 0, 0 }, { "ForeignName", JT_TEXT, 10, 255, 0 }, { "RmtInfoShort", JT_BINARY, 11, 255, 0 },
    { "RmtInfoLong", JT_OLE, 12, 0, 0 }, { "Lv", JT_OLE, 13, 0, 0 }, { "LvProp", JT_OLE, 14, 0, 0 },
    { "LvModule", JT_OLE, 15, 0, 0 }, { "LvExtra", JT_OLE, 16, 0, 0 }, { 0, 0, 0, 0, 0 } };
static const ccol C_ACE[] = { { "ObjectId", JT_LONG, 0, 4, 1 }, { "SID", JT_BINARY, 1, 255, 0 }, { "ACM", JT_LONG, 2, 4, 1 },
    { "FInheritable", JT_BOOL, 3, 1, 1 }, { 0, 0, 0, 0, 0 } };
static const ccol C_QRY[] = { { "ObjectId", JT_LONG, 0, 4, 1 }, { "Attribute", JT_BYTE, 1, 1, 1 }, { "Order", JT_BINARY, 2, 255, 0 },
    { "Name1", JT_TEXT, 3, 255, 0 }, { "Name2", JT_TEXT, 4, 255, 0 }, { "Expression", JT_MEMO, 5, 0, 0 },
    { "Flag", JT_INT, 6, 2, 1 }, { "LvExtra", JT_LONG, 7, 4, 1 }, { 0, 0, 0, 0, 0 } };    /* LvExtra: Jet 4 only */
static const ccol C_REL[] = { { "szRelationship", JT_TEXT, 0, 255, 0 }, { "grbit", JT_LONG, 1, 4, 1 }, { "ccolumn", JT_LONG, 2, 4, 1 },
    { "icolumn", JT_LONG, 3, 4, 1 }, { "szObject", JT_TEXT, 4, 255, 0 }, { "szColumn", JT_TEXT, 5, 255, 0 },
    { "szReferencedObject", JT_TEXT, 6, 255, 0 }, { "szReferencedColumn", JT_TEXT, 7, 255, 0 }, { 0, 0, 0, 0, 0 } };

typedef struct cidx { const char *name; const char *cols[3]; int f4, f3, prim; } cidx;
static const cidx X_OBJ[] = { { "ParentIdName", { "ParentId", "Name", 0 }, 0x81, 0x01, 0 }, { "Id", { "Id", 0, 0 }, 0x81, 0x01, 1 }, { 0, { 0 }, 0, 0, 0 } };
static const cidx X_ACE[] = { { "ObjectId", { "ObjectId", 0, 0 }, 0x88, 0x08, 0 }, { 0, { 0 }, 0, 0, 0 } };
static const cidx X_QRY[] = { { "ObjectIdAttribute", { "ObjectId", "Attribute", "Order" }, 0x81, 0x01, 1 }, { 0, { 0 }, 0, 0, 0 } };
static const cidx X_REL[] = { { "szRelationship", { "szRelationship", 0, 0 }, 0x82, 0x02, 0 }, { "szObject", { "szObject", 0, 0 }, 0x82, 0x02, 0 },
    { "szReferencedObject", { "szReferencedObject", 0, 0 }, 0x82, 0x02, 0 }, { 0, { 0 }, 0, 0, 0 } };
static const char *const LV_OBJ[] = { "Connect", "Database", "Lv", "LvExtra", "LvModule", "LvProp", "RmtInfoLong", 0 };
static const char *const LV_QRY[] = { "Expression", 0 };

static int name_cmp(const void *a, const void *b)
{
    const jetcol *x = a, *y = b;
    return jet_strcasecmp8(x->name, y->name);
}

static jtable *catalog_table(int jet4, const char *name, uint32_t page, const ccol *cols, const cidx *ix, const char *const *lv)
{
    jtable *t = jet_table_new(name);
    t->page = page;
    t->type = 0x53;
    int n = 0;
    while (cols[n].name) n++;
    if (!jet4 && cols == C_QRY) n = 7;          /* Jet 3: no LvExtra */
    t->cols = calloc((size_t)n, sizeof(jetcol));
    t->ncols = n;
    for (int i = 0; i < n; i++) {
        jetcol *c = &t->cols[i];
        c->name = strdup(cols[i].name);
        c->type = cols[i].type; c->num = cols[i].num;
        c->len = (cols[i].type == JT_TEXT || cols[i].type == JT_BINARY) && jet4 ? cols[i].len * 2 : cols[i].len;
        c->flags = (cols[i].fixed ? 0x13 : 0x12) | (!strcmp(c->name, "Owner") || !strcmp(c->name, "SID") ? 0x20 : 0);
        c->misc = 0x409;
        c->misc2 = jet4 ? 0 : 0x4E4;
    }
    if (jet4) qsort(t->cols, (size_t)n, sizeof(jetcol), name_cmp);     /* Jet 4 lists them by name */
    int nl = 0;
    while (lv && lv[nl]) nl++;
    t->lvorder = malloc(sizeof(int) * (size_t)(nl + 1));
    for (int i = 0; i < nl; i++) t->lvorder[i] = jet_col(t, lv[i])->num;
    t->nlv = nl;
    int nx = 0;
    while (ix[nx].name) nx++;
    t->ridx = calloc((size_t)nx, sizeof(jridx));
    t->lidx = calloc((size_t)nx, sizeof(jlidx));
    t->nridx = t->nlidx = nx;
    for (int i = 0; i < nx; i++) {
        jridx *r = &t->ridx[i];
        for (int k = 0; k < 3 && ix[i].cols[k]; k++) { r->col[k] = jet_col(t, ix[i].cols[k])->num; r->order[k] = 1; r->n++; }
        r->flags = jet4 ? ix[i].f4 : ix[i].f3;
        jlidx *l = &t->lidx[i];
        l->name = strdup(ix[i].name);
        l->num = i; l->real = i; l->type = ix[i].prim; l->relidx = 0xFFFFFFFFu; l->cup = l->cdel = 4;
    }
    t->width = n;
    return t;
}

static double vb_now_date(void)
{
    time_t now = time(0);
    struct tm tm;
#ifdef _WIN32
    localtime_s(&tm, &now);
#else
    localtime_r(&now, &tm);
#endif
    /* days from 1899-12-30 */
    long y = tm.tm_year + 1900, m = tm.tm_mon + 1, d = tm.tm_mday;
    long a = (14 - m) / 12, yy = y + 4800 - a, mm = m + 12 * a - 3;
    long jdn = d + (153 * mm + 2) / 5 + 365 * yy + yy / 4 - yy / 100 + yy / 400 - 32045;
    return (double)(jdn - 2415019) + (tm.tm_hour * 3600 + tm.tm_min * 60 + tm.tm_sec) / 86400.0;
}

static void set_text(jdb *db, jv *v, const char *s)
{
    uint16_t *w;
    int n = jet_utf8_to_utf16(s, &w);
    if (db->jet4) {
        unsigned char *b = malloc((size_t)n * 2 + 1);
        for (int i = 0; i < n; i++) { b[2 * i] = (unsigned char)w[i]; b[2 * i + 1] = (unsigned char)(w[i] >> 8); }
        jet_val_set(v, b, n * 2);
        free(b);
    } else {
        unsigned char *b = malloc((size_t)n * 2 + 1);
        int k = jet_utf16_to_cp(db->codepage, w, n, b);
        jet_val_set(v, b, k);
        free(b);
    }
    free(w);
}
static void set_u32(jv *v, uint32_t x) { unsigned char b[4]; P32(b, x); jet_val_set(v, b, 4); }

static void obj(jdb *db, uint32_t id, uint32_t parent, const char *name, int type, unsigned owner, double now)
{
    jtable *t = db->t[0];
    int r = jet_row_add(t);
    unsigned char b[8];
    set_u32(jet_val(t, r, jet_col(t, "Id")->num), id);
    set_u32(jet_val(t, r, jet_col(t, "ParentId")->num), parent);
    set_text(db, jet_val(t, r, jet_col(t, "Name")->num), name);
    P16(b, (unsigned)type); jet_val_set(jet_val(t, r, jet_col(t, "Type")->num), b, 2);
    memcpy(b, &now, 8);
    jet_val_set(jet_val(t, r, jet_col(t, "DateCreate")->num), b, 8);
    jet_val_set(jet_val(t, r, jet_col(t, "DateUpdate")->num), b, 8);
    jet_sid(db->jet4, owner, b);
    jet_val_set(jet_val(t, r, jet_col(t, "Owner")->num), b, 2);
    set_u32(jet_val(t, r, jet_col(t, "Flags")->num), 0x80000000u);
}

jdb *jet_new(int jet4, int codepage);
jdb *jet_new_order(int jet4, int codepage, int lcid);
jdb *jet_new_order(int jet4, int codepage, int lcid)
{
    /* Jet 4 in another sort order: its LCID, code page, and a date whose SID key is known (Jet 3: General) */
    cur_date = JET4_DATE; cur_key = JET4_KEY;
    int ocp = 0;
    if (jet4 && lcid && lcid != 0x409)
        for (unsigned i = 0; i < sizeof JET4_ORDERS / sizeof JET4_ORDERS[0]; i++)
            if (JET4_ORDERS[i].lcid == lcid) { cur_date = JET4_ORDERS[i].date; cur_key = JET4_ORDERS[i].key; ocp = JET4_ORDERS[i].cp; break; }
    jdb *db = jet_new(jet4, codepage);
    if (ocp) {
        P32(db->page0 + 0x6E, (uint32_t)lcid);
        P16(db->page0 + 0x3C, (unsigned)ocp);
        for (int i = 0; i < db->nt; i++)
            for (int k = 0; k < db->t[i]->ncols; k++)
                if ((db->t[i]->cols[k].type == JT_TEXT || db->t[i]->cols[k].type == JT_MEMO) && db->t[i]->cols[k].misc == 0x409)
                    db->t[i]->cols[k].misc = lcid;     /* (as Jet: c_0408 MSysObjects.Name, Connect 0x408) */
    }
    cur_date = JET4_DATE; cur_key = JET4_KEY;
    return db;
}

jdb *jet_new(int jet4, int codepage)
{
    jdb *db = calloc(1, sizeof *db);
    db->jet4 = jet4;
    db->ps = jet4 ? 4096 : 2048;
    db->codepage = codepage ? codepage : 1252;
    db->page0 = page0(jet4);
    db->t = malloc(sizeof(jtable *) * 16);
    db->t[db->nt++] = catalog_table(jet4, "MSysObjects", 2, C_OBJ, X_OBJ, LV_OBJ);
    db->t[db->nt++] = catalog_table(jet4, "MSysACEs", 3, C_ACE, X_ACE, 0);
    db->t[db->nt++] = catalog_table(jet4, "MSysQueries", 4, C_QRY, X_QRY, LV_QRY);
    db->t[db->nt++] = catalog_table(jet4, "MSysRelationships", 5, C_REL, X_REL, 0);
    double now = vb_now_date();
    obj(db, ID_TABLES, 0x0F000000u, "Tables", 3, SID_SYSTEM, now);
    obj(db, ID_DATABASES, 0x0F000000u, "Databases", 3, SID_SYSTEM, now);
    obj(db, ID_RELATIONSHIPS, 0x0F000000u, "Relationships", 3, SID_SYSTEM, now);
    obj(db, ID_MSYSDB, ID_DATABASES, "MSysDb", 2, SID_ADMIN, now);
    static const char *const names[4] = { "MSysObjects", "MSysACEs", "MSysQueries", "MSysRelationships" };
    for (int i = 0; i < 4; i++) obj(db, 2u + (uint32_t)i, ID_TABLES, names[i], 1, SID_SYSTEM, now);
    static const struct { uint32_t oid; unsigned sid; uint32_t acm; int inh; } A[] = {
        { 2, SID_ADMIN, 0x60000, 0 }, { 3, SID_ADMIN, 0x60000, 0 }, { 4, SID_ADMIN, 0x60000, 0 }, { 5, SID_ADMIN, 0xE0000, 0 },
        { ID_TABLES, SID_CREATOR, 0xF00FE, 1 }, { ID_TABLES, SID_ADMIN, 0x60001, 0 },
        { ID_RELATIONSHIPS, SID_CREATOR, 0xF00FE, 1 }, { ID_RELATIONSHIPS, SID_ADMIN, 0x60001, 0 },
        { ID_DATABASES, SID_ADMIN, 0x60000, 0 }, { ID_MSYSDB, SID_ADMIN, 0x6000E, 0 }, { ID_MSYSDB, SID_ADMINS, 0xE, 0 },
        { 4, SID_ADMINS, 0x14, 0 }, { 5, SID_ADMINS, 0x14, 0 }, { 2, SID_ADMINS, 0x14, 0 },
        { ID_TABLES, SID_ADMINS, 0xFFEFF, 1 }, { ID_RELATIONSHIPS, SID_ADMINS, 0xFFFFF, 1 } };
    jtable *a = db->t[1];
    for (unsigned i = 0; i < sizeof A / sizeof A[0]; i++) {
        int r = jet_row_add(a);
        unsigned char s[2];
        set_u32(jet_val(a, r, jet_col(a, "ObjectId")->num), A[i].oid);
        jet_sid(jet4, A[i].sid, s);
        jet_val_set(jet_val(a, r, jet_col(a, "SID")->num), s, 2);
        set_u32(jet_val(a, r, jet_col(a, "ACM")->num), A[i].acm);
        jet_val(a, r, jet_col(a, "FInheritable")->num)->n = A[i].inh;
    }
    return db;
}
