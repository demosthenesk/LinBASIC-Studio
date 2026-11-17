/* resource.c - LoadResString, LoadResData, LoadResPicture: the resource file of the project
 * (ResFile32= of the .vbp), which lbvc puts into the program as a table.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include <strings.h>

static const vb_res *res_tab;
static vb_i32 res_n;

void vb_res_init(const vb_res *table, vb_i32 n) { res_tab = table; res_n = n; }

vbobj (*vb_picture_data_hook)(const char *data, vb_i32 n, vb_i32 type);

static int same_name(const char *a, vbs b) { return strcasecmp(a, vb_stoutf8(b)) == 0; }

/* type and name: a number (id) or a name */
static const vb_res *find(vb_i32 type_id, vbs type_name, vb_i32 id, vbs name)
{
    for (vb_i32 i = 0; i < res_n; i++) {
        const vb_res *r = &res_tab[i];
        if (type_name.p ? !(r->type_name && same_name(r->type_name, type_name)) : r->type_id != type_id) continue;
        if (name.p ? !(r->name && same_name(r->name, name)) : (r->name || r->id != id)) continue;
        return r;
    }
    return 0;
}

static void not_found(vbvar id)
{
    vb_i32 m = vb_mark();
    vbs t = vb_v2s(id);
    char *s = vb_stoutf8(t);
    char buf[300];
    snprintf(buf, sizeof buf, "Resource with identifier '%s' not found", s);
    vb_raise_full(326, vb_stmp(0), vb_sfromutf8(buf));
    vb_release(m);
}

/* id of LoadResData/LoadResPicture: a number, or a string that is the name */
static void id_of(vbvar v, vb_i32 *id, vbs *name)
{
    *id = 0; name->p = 0; name->len = 0;
    if (vb_vartype(v) == VT_BSTR) *name = vb_v2s(v);
    else *id = vb_v2l(v);
}

vbs vb_loadresstring(vb_i32 id)
{
    vbs none = { 0, 0, 0 };
    const vb_res *r = find(6, none, (id >> 4) + 1, none);
    if (r) {                                    /* a block of 16 strings: u16 length, UTF-16 */
        const unsigned char *p = r->data, *end = r->data + r->size;
        for (int i = 0; i < 16 && p + 2 <= end; i++) {
            int len = p[0] | p[1] << 8;
            p += 2;
            if (i == (id & 15)) {
                if (len == 0) break;
                vbs s = vb_stmp(len);
                for (int k = 0; k < len && p + 2 * k + 1 < end; k++) s.p[k] = (vb_wchar)(p[2 * k] | p[2 * k + 1] << 8);
                return s;
            }
            p += 2 * len;
        }
    }
    not_found(vb_v_l(id));
    return vb_stmp(0);
}

vbvar vb_loadresdata(vbvar idv, vbvar typev)
{
    vb_i32 id, tid; vbs name, tname;
    id_of(idv, &id, &name);
    id_of(typev, &tid, &tname);
    const vb_res *r = find(tid, tname, id, name);
    if (!r) { not_found(idv); return vb_v_empty(); }
    vbarr *a = vb_atmp();
    if (r->size) {
        vb_adim(a, 1, VBA_PLAIN, 0, 1, 0, r->size - 1);
        if (a->data) memcpy(a->data, r->data, r->size);
    } else vb_aempty(a, 1, VBA_PLAIN);
    vbvar v = vb_v_empty();
    v.vt = VT_ARRAY | VT_UI1; v.u.a = a;
    return v;
}

static void put16(unsigned char *p, unsigned v) { p[0] = v & 0xFF; p[1] = (v >> 8) & 0xFF; }
static void put32(unsigned char *p, unsigned v) { put16(p, v & 0xFFFF); put16(p + 2, v >> 16); }
static unsigned get16(const unsigned char *p) { return p[0] | p[1] << 8; }
static unsigned get32(const unsigned char *p) { return get16(p) | get16(p + 2) << 16; }

/* RT_BITMAP has no BITMAPFILEHEADER: the 14 bytes of a .bmp file before it */
static vbobj bitmap(const vb_res *r)
{
    if (r->size < 40) return 0;
    const unsigned char *d = r->data;
    unsigned hsz = get32(d), bits = get16(d + 14), comp = get32(d + 16), used = get32(d + 32);
    unsigned colors = used ? used : bits <= 8 ? 1u << bits : 0;
    unsigned off = 14 + hsz + colors * 4 + (comp == 3 && hsz == 40 ? 12 : 0);
    unsigned n = 14 + r->size;
    unsigned char *f = (unsigned char *)malloc(n);
    f[0] = 'B'; f[1] = 'M'; put32(f + 2, n); put32(f + 6, 0); put32(f + 10, off);
    memcpy(f + 14, d, r->size);
    vbobj o = vb_picture_data_hook((const char *)f, n, 1);
    free(f);
    return o;
}

/* RT_GROUP_ICON / RT_GROUP_CURSOR and their RT_ICON / RT_CURSOR images: an .ico file */
static vbobj icon(const vb_res *g, int cursor)
{
    if (g->size < 6) return 0;
    const unsigned char *d = g->data;
    unsigned cnt = get16(d + 4);
    if (!cnt || g->size < 6 + 14 * cnt) return 0;
    const unsigned char *e = d + 6;                /* the first (VB6 uses the best size; one is enough) */
    vbs none = { 0, 0, 0 };
    const vb_res *img = find(cursor ? 1 : 3, none, get16(e + 12), none);
    if (!img) return 0;
    const unsigned char *dib = img->data;
    unsigned dsz = img->size;
    if (cursor) { dib += 4; dsz -= 4; }           /* the hot spot */
    unsigned w = get32(dib + 4), h = get32(dib + 8) / 2;
    unsigned n = 6 + 16 + dsz;
    unsigned char *f = (unsigned char *)calloc(1, n);
    put16(f + 2, 1); put16(f + 4, 1);
    f[6] = w >= 256 ? 0 : (unsigned char)w; f[7] = h >= 256 ? 0 : (unsigned char)h;
    put16(f + 10, 1); put16(f + 12, get16(dib + 14));
    put32(f + 14, dsz); put32(f + 18, 22);
    memcpy(f + 22, dib, dsz);
    vbobj o = vb_picture_data_hook((const char *)f, n, 3);
    free(f);
    return o;
}

vbobj vb_loadrespicture(vbvar idv, vb_i32 format)
{
    vb_i32 id; vbs name;
    id_of(idv, &id, &name);
    vbs none = { 0, 0, 0 };
    if (!vb_picture_data_hook) { vb_raise(481); return 0; }
    const vb_res *r = find(format == 0 ? 2 : format == 1 ? 14 : 12, none, id, name);
    if (!r) { not_found(idv); return 0; }
    vbobj o = format == 0 ? bitmap(r) : icon(r, format == 2);
    if (!o) vb_raise(481);                       /* Invalid picture */
    return o;
}
