/* stdfmt.c - the Microsoft Data Formatting Object Library (MSSTDFMT.DLL, StdFormat) of LinBASIC Studio
 * (phase 7.7): StdDataFormat, StdDataValue and StdDataFormats, and what the bound controls do with a
 * StdDataFormat: a value of a field into what the control shows (vb_dfmt_format) and back (vb_dfmt_unformat).
 *
 * As VB6 does it (VB6 oracles f_dfmt, f_dfmt2):
 * - Type 0 General: the value as text; 1 Custom: Format$(value, Format); 4 Checkbox: 1 / 0; 5 Boolean: TrueValue /
 *   FalseValue; 2 Picture, 3 Object, 6 Bytes: the value itself.
 * - Null: "" with General, NullValue with the other types (back: the text NullValue is Null again).
 * - The Format of the .frm is written in the language of its LCID ("0.00" with 1032, Greek, groups the digits);
 *   the text the control shows follows the regional settings of the user.
 * - Format event: DataValue.Value is the formatted value; UnFormat: the text of the control. Both can change it.
 * - Changed event: each property set; Format = makes the Type Custom (two Changed when it was not).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"

enum { fmtGeneral = 0, fmtCustom = 1, fmtPicture = 2, fmtObject = 3, fmtCheckbox = 4, fmtBoolean = 5, fmtBytes = 6 };

typedef struct dfmt {
    vbobjh h;
    vb_i32 type, day, week, lcid;       /* lcid: the language the Format is written in (0: the user's) */
    vbs fmt;
    vbvar tv, fv, nv;                   /* TrueValue, FalseValue, NullValue */
} dfmt;

typedef struct dval {
    vbobjh h;
    vbvar v;
    vbobj target;                       /* counted */
} dval;

typedef struct dfmts {
    vbobjh h;
    vb_i32 n, cap;
    vbobj *f;                           /* counted */
    vbs *key;                           /* the Index of Add (a field name) */
} dfmts;

extern const vb_class vb_cls_stddataformat, vb_cls_stddatavalue, vb_cls_stddataformats;

static vbvar ret_obj(vbobj o) { return vb_vtmp(vb_v_obj(o)); }

/* ---------------------------------------------------------------- events */
static void fire(vbobj src, const char *ev, int argc, vbvar *argv)
{
    if (!src || !src->sinks) return;
    vbobj *ls; const vb_fn **ts;
    vb_i32 n = vb_sinks_begin(src, &ls, &ts);
    for (vb_i32 k = 0; k < n && !vb_errpend; k++)
        for (const vb_member *mm = (const vb_member *)ts[k]; mm && mm->name; mm++)
            if (!strcmp(mm->name, ev)) { mm->fn(ls[k], argc, argv); break; }
    vb_sinks_end(ls, ts, n);
}
static void changed(dfmt *f) { fire((vbobj)f, "changed", 0, 0); }

/* the Format or UnFormat event: DataValue.Value = *v, then what the program left in it */
static void value_event(dfmt *f, const char *ev, vbvar *v, vbobj target)
{
    if (!f->h.sinks) return;
    vb_i32 m = vb_mark();
    dval *d = (dval *)vb_onew(&vb_cls_stddatavalue);
    vb_oaddref((vbobj)d);
    vb_vset(&d->v, *v);
    if (target) { d->target = vb_oid(target); vb_oaddref(d->target); }
    vbvar a = vb_v_obj((vbobj)d);
    fire((vbobj)f, ev, 1, &a);
    vb_release(m);
    *v = vb_vtmp(d->v);
    vb_orelease((vbobj)d);
}

/* ---------------------------------------------------------------- StdDataFormat */
static void fmt_init(vbobj o)
{
    dfmt *f = (dfmt *)o;
    f->type = fmtGeneral;
    vb_vset(&f->tv, vb_v_bool(-1));
    vb_vset(&f->fv, vb_v_bool(0));
    vb_vset(&f->nv, vb_v_s(vb_stmp(0)));
}
static void fmt_clear(vbobj o)
{
    dfmt *f = (dfmt *)o;
    vb_sfree(&f->fmt);
    vb_vfree(&f->tv); vb_vfree(&f->fv); vb_vfree(&f->nv);
}

vbobj vb_dfmt_new(void)
{
    vbobj o = vb_onew(&vb_cls_stddataformat);
    vb_oaddref(o);
    return o;
}

/* the DataFormat of the .frm: tfn = HaveTrueFalseNull (TrueValue, FalseValue, NullValue are there) */
void vb_dfmt_load(vbobj o, vb_i32 type, vbs fmt, vb_i32 tfn, vbvar tv, vbvar fv, vbvar nv, vb_i32 day, vb_i32 week, vb_i32 lcid)
{
    dfmt *f = (dfmt *)vb_oid(o);
    if (!f || f->h.cls != &vb_cls_stddataformat) return;
    f->type = type >= 0 && type <= 6 ? type : 0;
    vb_sset(&f->fmt, fmt);
    if (tfn) { vb_vset(&f->tv, tv); vb_vset(&f->fv, fv); vb_vset(&f->nv, nv); }
    f->day = day >= 0 && day <= 7 ? day : 0;
    f->week = week >= 0 && week <= 3 ? week : 0;
    f->lcid = lcid;
}

vb_bool vb_is_dfmt(vbobj o) { o = vb_oid(o); return o && o->cls == &vb_cls_stddataformat; }
vb_i32 vb_dfmt_kind(vbobj o)                    /* its Type; -1: no StdDataFormat */
{
    dfmt *f = (dfmt *)vb_oid(o);
    return f && f->h.cls == &vb_cls_stddataformat ? f->type : -1;
}

/* a number Format written in the language of lcid, in the form Format$ reads (the "." decimal point, the ","
 * thousands): MSSTDFMT reads it with that language (VB6 oracle f_dfmt: "0.00" with 1032 -> 1.235). A date or
 * time Format is left as it is. */
static int date_letter(vb_wchar c)
{
    c = vb_wlower(c);
    return c == 'd' || c == 'm' || c == 'y' || c == 'h' || c == 'n' || c == 's' || c == 'w' || c == 'q' || c == 'c' ||
           c == 't' || c == 'a';
}
static vbs lcid_format(vbs s, vb_i32 lcid)
{
    char dec = '.', grp = ',';
    if (lcid == 0) return s;                    /* a format of the program: the notation of Format$ (VB6 oracle f_dfmt) */
    if (lcid == 1024 || lcid == 2048 || lcid == 3072) { dec = vb_lc_dec; grp = vb_lc_grp; }
    else if (!vb_lc_seps(lcid, &dec, &grp)) return s;
    if (dec == '.' && grp == ',') return s;
    int digits = 0;
    for (vb_i32 i = 0; i < s.len; i++) {        /* only a number format: no date letters outside the literals */
        vb_wchar c = s.p[i];
        if (c == '\\') { i++; continue; }
        if (c == '"') { while (++i < s.len && s.p[i] != '"') ; continue; }
        if (c == '0' || c == '#') digits = 1;
        else if (date_letter(c)) return s;
    }
    if (!digits) return s;
    vbs r = vb_stmp(s.len);
    for (vb_i32 i = 0; i < s.len; i++) {
        vb_wchar c = s.p[i];
        r.p[i] = c;
        if (c == '\\') { if (i + 1 < s.len) { r.p[i + 1] = s.p[i + 1]; i++; } continue; }
        if (c == '"') { while (++i < s.len) { r.p[i] = s.p[i]; if (s.p[i] == '"') break; } continue; }
        if (c == (vb_wchar)(unsigned char)dec) r.p[i] = '.';
        else if (c == (vb_wchar)(unsigned char)grp && (grp == '.' || grp == ',')) r.p[i] = ',';
        else if ((c == ' ' || c == 0xA0) && grp == ' ' && i > 0 && i + 1 < s.len &&
                 (s.p[i - 1] == '#' || s.p[i - 1] == '0') && (s.p[i + 1] == '#' || s.p[i + 1] == '0')) r.p[i] = ',';
    }
    return r;
}

/* a value of a field -> what the bound control shows (the Format event can change it); a temporary */
vbvar vb_dfmt_format(vbobj o, vbvar v, vbobj target)
{
    dfmt *f = (dfmt *)vb_oid(o);
    if (!f || f->h.cls != &vb_cls_stddataformat) return v;
    vb_i32 save = vb_errpend;
    vbvar r;
    if (v.vt == VT_DISPATCH && f->type != fmtObject) { v = vb_late(v, "", VBM_GET, 0, 0); if (vb_errpend && !save) { vb_err_clear_pending(); v = vb_v_null(); } }
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) {
        if (f->type == fmtGeneral) r = vb_v_s(vb_stmp(0));
        else if (f->type == fmtPicture || f->type == fmtObject || f->type == fmtBytes) r = v;
        else r = vb_vtmp(f->nv);
    } else {
        switch (f->type) {
        case fmtCustom:
            if (f->fmt.len) { r = vb_v_s(vb_format(v, lcid_format(f->fmt, f->lcid), 1, f->day, f->week)); break; }
            /* fall through: Custom without a Format */
        case fmtGeneral:
            r = vb_v_s(vb_v2s(v));
            break;
        case fmtCheckbox: {
            vb_bool b = vb_v2bool(v);
            r = vb_v_i(b ? 1 : 0);
            break; }
        case fmtBoolean: {
            vb_bool b = vb_v2bool(v);
            r = vb_vtmp(b ? f->tv : f->fv);
            break; }
        default:
            r = v;
        }
        if (vb_errpend && !save) {              /* not of the format (a text as Boolean): the value as text */
            vb_err_clear_pending();
            r = vb_v_s(vb_v2s(v));
            if (vb_errpend && !save) { vb_err_clear_pending(); r = v; }
        }
    }
    value_event(f, "format", &r, target);
    return r;
}

static int same_text(vbvar v, vbvar w)
{
    if (w.vt == VT_NULL || w.vt == VT_EMPTY || v.vt != VT_BSTR) return 0;
    vb_i32 save = vb_errpend;
    vbs b = vb_v2s(w);
    if (vb_errpend && !save) { vb_err_clear_pending(); return 0; }
    return vb_scmp(v.u.s, b, 1) == 0;
}

/* what the bound control holds -> the value for the field (UnFormat event first); a String still has to become
 * the type of the field (vb_db_strto); -1: it does not fit the format */
vb_i32 vb_dfmt_unformat(vbobj o, vbvar v, vbobj target, vbvar *out)
{
    dfmt *f = (dfmt *)vb_oid(o);
    *out = v;
    if (!f || f->h.cls != &vb_cls_stddataformat) return 0;
    value_event(f, "unformat", &v, target);
    *out = v;
    if (vb_errpend) return -1;
    if (v.vt != VT_BSTR) return 0;              /* the program gave a value of its own */
    switch (f->type) {
    case fmtBoolean:
        if (same_text(v, f->tv)) { *out = vb_v_bool(-1); return 0; }
        if (same_text(v, f->fv)) { *out = vb_v_bool(0); return 0; }
        if (same_text(v, f->nv) || (v.u.s.len == 0 && f->nv.vt == VT_NULL)) { *out = vb_v_null(); return 0; }
        return -1;
    case fmtCheckbox: {
        if (v.u.s.len == 0 || same_text(v, f->nv)) { *out = vb_v_null(); return 0; }
        vb_i32 save = vb_errpend;
        double d = vb_s2d(v.u.s);
        if (vb_errpend && !save) { vb_err_clear_pending(); return -1; }
        *out = d == 2 ? vb_v_null() : vb_v_bool(d != 0 ? -1 : 0);
        return 0; }
    case fmtCustom:
        if (same_text(v, f->nv)) *out = vb_v_null();
        return 0;
    default:
        return 0;
    }
}

/* a String into the type of an ADO field (DataTypeEnum), with the regional settings as CDbl / CDate read it;
 * "" is Null in a field that is no text; -1 when it does not fit (*out: the value as it was) */
vb_i32 vb_db_strto(vbvar v, vb_i32 t, vbvar *out)
{
    *out = v;
    if (v.vt != VT_BSTR) return 0;
    int text = t == 8 || t == 129 || t == 130 || t == 200 || t == 201 || t == 202 || t == 203 || t == 72 ||
               t == 128 || t == 204 || t == 205 || t == 0;
    if (text) return 0;
    if (v.u.s.len == 0) { *out = vb_v_null(); return 0; }
    vb_i32 save = vb_errpend;
    vbvar r = v;
    switch (t) {
    case 2: case 16: r = vb_v_i(vb_v2i(v)); break;                       /* adSmallInt, adTinyInt */
    case 17: r = vb_v_byte(vb_v2byte(v)); break;                          /* adUnsignedTinyInt */
    case 3: case 18: case 19: r = vb_v_l(vb_v2l(v)); break;              /* adInteger ... */
    case 4: r = vb_v_f(vb_v2f(v)); break;                                 /* adSingle */
    case 5: case 20: case 21: r = vb_v_d(vb_v2d(v)); break;              /* adDouble, adBigInt */
    case 6: r = vb_v_c(vb_v2c(v)); break;                                 /* adCurrency */
    case 7: case 133: case 134: case 135: case 64: r = vb_v_date(vb_v2date(v)); break;
    case 11: r = vb_v_bool(vb_v2bool(v)); break;                          /* adBoolean */
    case 14: case 131: case 139: r = vb_cdec(v); break;                   /* adDecimal, adNumeric */
    default: return 0;
    }
    if (vb_errpend && !save) { vb_err_clear_pending(); return -1; }
    *out = r;
    return 0;
}

static vbvar fmt_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dfmt *f = (dfmt *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!strcmp(name, "type")) {
        if (!set) return vb_v_l(f->type);
        vb_i32 n = vb_v2l(val);
        if (vb_errpend) return vb_v_empty();
        if (n < 0 || n > 6) { vb_raise(5); return vb_v_empty(); }
        f->type = n;
        changed(f);
        return vb_v_empty();
    }
    if (!strcmp(name, "format")) {
        if (!set) return vb_v_s(vb_scopy(f->fmt));
        vbs s = vb_v2s(val);
        if (vb_errpend) return vb_v_empty();
        vb_sset(&f->fmt, s);
        int was = f->type;
        f->type = fmtCustom;                    /* Format = makes the format Custom (VB6 oracle f_dfmt) */
        changed(f);
        if (was != fmtCustom && !vb_errpend) changed(f);
        return vb_v_empty();
    }
    vbvar *vp = !strcmp(name, "truevalue") ? &f->tv : !strcmp(name, "falsevalue") ? &f->fv :
                !strcmp(name, "nullvalue") ? &f->nv : 0;
    if (vp) {
        if (!set) return vb_vtmp(*vp);
        vb_vset(vp, val);
        changed(f);
        return vb_v_empty();
    }
    if (!strcmp(name, "firstdayofweek") || !strcmp(name, "firstweekofyear")) {
        int day = name[5] == 'd';
        vb_i32 *ip = day ? &f->day : &f->week;
        if (!set) return vb_v_l(*ip);
        vb_i32 n = vb_v2l(val);
        if (vb_errpend) return vb_v_empty();
        if (n < 0 || n > (day ? 7 : 3)) { vb_raise(5); return vb_v_empty(); }
        *ip = n;
        changed(f);
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- StdDataValue */
static void val_clear(vbobj o)
{
    dval *d = (dval *)o;
    vb_vfree(&d->v);
    if (d->target) vb_orelease(d->target);
}
static vbvar val_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dval *d = (dval *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!strcmp(name, "value") || !*name) {
        if (!set) return vb_vtmp(d->v);
        vb_vset(&d->v, val);
        return vb_v_empty();
    }
    if (!strcmp(name, "targetobject")) {
        if (!set) return ret_obj(d->target);
        vbobj n = vb_v2obj(val);
        if (vb_errpend) return vb_v_empty();
        n = vb_oid(n);
        if (n) vb_oaddref(n);
        if (d->target) vb_orelease(d->target);
        d->target = n;
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- StdDataFormats */
static void fmts_clear_all(dfmts *c)
{
    for (vb_i32 i = 0; i < c->n; i++) { vb_orelease(c->f[i]); vb_sfree(&c->key[i]); }
    c->n = 0;
}
static void fmts_clear(vbobj o)
{
    dfmts *c = (dfmts *)o;
    fmts_clear_all(c);
    free(c->f); free(c->key);
}
/* Item / Remove: a number is a position from 0, a text the Index of Add (VB6 oracle f_dfmt); -1: none */
static vb_i32 fmts_find(dfmts *c, vbvar k)
{
    if (k.vt == VT_BSTR) {
        for (vb_i32 i = 0; i < c->n; i++) if (vb_scmp(c->key[i], k.u.s, 1) == 0) return i;
        return -1;
    }
    vb_i32 save = vb_errpend;
    vb_i32 i = vb_v2l(k);
    if (vb_errpend && !save) { vb_err_clear_pending(); return -1; }
    return i >= 0 && i < c->n ? i : -1;
}
static vb_i32 fmts_each_count(vbobj o) { return ((dfmts *)vb_oid(o))->n; }
static vbvar fmts_each_item(vbobj o, vb_i32 i)
{
    dfmts *c = (dfmts *)vb_oid(o);
    return i >= 0 && i < c->n ? ret_obj(c->f[i]) : vb_v_empty();
}
static vbvar fmts_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    dfmts *c = (dfmts *)vb_oid(o);
    *found = 1;
    if (!strcmp(name, "count")) return vb_v_l(c->n);
    if (!strcmp(name, "item") || !*name) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        vb_i32 i = fmts_find(c, argv[0]);
        if (i < 0) { vb_raise(5); return vb_v_empty(); }
        return ret_obj(c->f[i]);
    }
    if (!strcmp(name, "add")) {
        if (argc < 2 || vb_ismissing(argv[1])) { vb_raise(13); return vb_v_empty(); }   /* without an Index: 13 */
        vbobj f = vb_v2obj(argv[0]);
        if (vb_errpend) return vb_v_empty();
        vbs k = vb_v2s(argv[1]);
        if (vb_errpend) return vb_v_empty();
        if (!f) { vb_raise(91); return vb_v_empty(); }
        if (c->n == c->cap) {
            c->cap = c->cap ? c->cap * 2 : 4;
            c->f = (vbobj *)realloc(c->f, sizeof(vbobj) * c->cap);
            c->key = (vbs *)realloc(c->key, sizeof(vbs) * c->cap);
        }
        f = vb_oid(f);
        vb_oaddref(f);
        c->f[c->n] = f;
        memset(&c->key[c->n], 0, sizeof(vbs));
        vb_sset(&c->key[c->n], k);
        c->n++;                                 /* the same Index twice is no error (VB6) */
        return vb_v_empty();
    }
    if (!strcmp(name, "remove")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        vb_i32 i = fmts_find(c, argv[0]);
        if (i < 0) { vb_raise(5); return vb_v_empty(); }
        vb_orelease(c->f[i]);
        vb_sfree(&c->key[i]);
        memmove(&c->f[i], &c->f[i + 1], sizeof(vbobj) * (c->n - i - 1));
        memmove(&c->key[i], &c->key[i + 1], sizeof(vbs) * (c->n - i - 1));
        c->n--;
        return vb_v_empty();
    }
    if (!strcmp(name, "clear")) { fmts_clear_all(c); return vb_v_empty(); }
    if (!strcmp(name, "_itemfield")) {          /* the Index of an item */
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        vb_i32 i = fmts_find(c, argv[0]);
        if (i < 0) { vb_raise(5); return vb_v_empty(); }
        return vb_v_s(vb_scopy(c->key[i]));
    }
    *found = 0;
    return vb_v_empty();
}

#define FB(f) 0, 0, 0, f
/* TypeName as VB6 gives it (VB6 oracle f_dfmt): the names of the interfaces */
const vb_class vb_cls_stddataformat = { "IStdDataFormatDisp", sizeof(dfmt), fmt_init, 0, fmt_clear, 0, 0, 0, FB(fmt_member) };
const vb_class vb_cls_stddatavalue = { "IStdDataValueDisp", sizeof(dval), 0, 0, val_clear, 0, 0, "value", FB(val_member) };
const vb_class vb_cls_stddataformats = { "IStdDataFormatsDisp", sizeof(dfmts), 0, 0, fmts_clear, 0, 0, "item",
                                         fmts_each_count, fmts_each_item, 0, fmts_member };

