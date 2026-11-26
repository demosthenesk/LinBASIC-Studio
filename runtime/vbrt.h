/*
 * vbrt.h - libvbrt, the runtime of programs compiled by lbvc (LinBASIC Studio).
 *
 * The interface uses no system headers, so generated programs can also be
 * compiled by the bundled Tiny C Compiler.
 *
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION).
 */
#ifndef VBRT_H
#define VBRT_H

/* the program compiled by tcc on Windows links the runtime DLLs (vbrt.dll ...): their variables are
 * reached through the import table */
#if defined(_WIN32) && defined(VB_DLL_IMPORT)
#define VB_IMPORT __declspec(dllimport)
#else
#define VB_IMPORT
#endif

typedef signed char vb_i8;
typedef unsigned char vb_u8;
typedef short vb_i16;
typedef unsigned short vb_u16;
typedef int vb_i32;
typedef unsigned int vb_u32;
typedef long long vb_i64;
typedef unsigned long long vb_u64;
typedef vb_i16 vb_bool;            /* True = -1, False = 0 */
typedef vb_i64 vb_cy;              /* Currency: value * 10000 */
typedef double vb_date;            /* Date: days since 30 Dec 1899 */
typedef vb_u16 vb_wchar;           /* strings are UTF-16, like VB6 BSTRs */

/* ---- strings -----------------------------------------------------------
 * A vbs is a string value.  Temporaries live until vb_release(mark) at the end
 * of the statement that made them; variables own their buffer (flags & 1). */
typedef struct vbs { vb_wchar *p; vb_i32 len; vb_i32 flags; } vbs;
/* flags: 1 = the buffer is owned (a variable); VBS_ODD = an odd number of bytes (ChrB, LeftB, StrConv ...):
   the last byte is in the low byte of p[len], after the len whole characters (Len does not count it) */
#define VBS_ODD 2

vbs vb_slit(const vb_wchar *s, vb_i32 len);   /* view of a constant      */
vbs vb_stmp(vb_i32 len);                       /* new temporary            */
vb_i32 vb_mark(void);
void vb_release(vb_i32 mark);
void vb_sset(vbs *dst, vbs src);               /* variable = value         */
void vb_sfree(vbs *v);
vbs vb_scat(vbs a, vbs b);                     /* a & b                    */
vbs vb_scopy(vbs s);                           /* temporary copy           */
vb_i32 vb_scmp(vbs a, vbs b, vb_i32 text);     /* <0 0 >0; text = Option Compare Text */
vbs vb_sfromutf8(const char *s);
void vb_sfix_init(vbs *v, vb_i32 n);           /* Dim s As String * n */
void vb_sset_fix(vbs *dst, vbs src, vb_i32 n); /* assignment to a String * n: pad / cut */
char *vb_stoutf8(vbs s);                       /* temporary C string       */

/* ---- errors ------------------------------------------------------------ */
extern VB_IMPORT vb_i32 vb_errpend;                      /* error raised, not handled yet */
extern VB_IMPORT char *vb_stack_lim;                   /* below it: Out of stack space (28) */
extern VB_IMPORT vb_i32 vb_fpchecks;     /* 0: Remove Floating Point Error Checks (native code) */
extern VB_IMPORT vb_i32 vb_fpstore;      /* the pending error is a floating-point one: the result is stored */
void vb_fpraise(vb_i32 number);          /* error 6 / 11 of a floating-point operation (none without checks) */
void vb_raise(vb_i32 number);                  /* run-time error with the standard text */
void vb_raise_arg(vb_i32 number, const char *arg);   /* ... whose |1 is arg */
void vb_raise_full(vb_i32 number, vbs source, vbs description);
void vb_err_raise(vb_i32 number, vbs source, vbs description, vb_i32 ctx, vb_i32 flags);   /* Err.Raise, Error */
void vb_err_raise_hf(vb_i32 number, vbs source, vbs description, vbs helpfile, vb_i32 ctx, vb_i32 flags);   /* with HelpFile (flag 8) */
vb_i32 vb_err_helpcontext(void);
void vb_err_set_helpcontext(vb_i32 n);
vbs vb_err_helpfile(void);
void vb_err_set_helpfile(vbs s);
void vb_err_take(vb_i32 erl);                  /* the pending error goes into Err */
void vb_err_clear(void);
vb_i32 vb_err_number(void);
vbs vb_err_description(void);
vbs vb_err_source(void);
void vb_err_set_number(vb_i32 n);
void vb_err_set_description(vbs s);
void vb_err_set_source(vbs s);
vb_i32 vb_erl(void);
void vb_unhandled(void);                       /* report and stop the program */

/* ---- numbers ----------------------------------------------------------- */
vb_i16 vb_add_i(vb_i32 a, vb_i32 b);
vb_i16 vb_sub_i(vb_i32 a, vb_i32 b);
vb_i16 vb_mul_i(vb_i32 a, vb_i32 b);
vb_i32 vb_add_l(vb_i64 a, vb_i64 b);
vb_i32 vb_sub_l(vb_i64 a, vb_i64 b);
vb_i32 vb_mul_l(vb_i64 a, vb_i64 b);
vb_u8 vb_add_b(vb_i32 a, vb_i32 b);
vb_u8 vb_sub_b(vb_i32 a, vb_i32 b);
vb_u8 vb_mul_b(vb_i32 a, vb_i32 b);
vb_cy vb_add_c(vb_cy a, vb_cy b);
vb_cy vb_sub_c(vb_cy a, vb_cy b);
vb_cy vb_mul_c(vb_cy a, vb_cy b);
double vb_div(double a, double b);
float vb_div_f(float a, float b);
vb_i32 vb_idiv(vb_i32 a, vb_i32 b);            /* \ after rounding the operands */
vb_i32 vb_mod(vb_i32 a, vb_i32 b);
double vb_pow(double a, double b);
double vb_chkd(double v);                      /* overflow check of a Double result */
/* Single and Double expressions: VB6 computes them on the x87 FPU in extended precision and rounds a value when
   it stores, passes or compares it (Allow Unrounded Floating Point Operations: not when it compares). vb_ld is
   that precision: 80 bits with gcc and mingw, a double with tcc on Windows. These take ONE vb_ld argument:
   the Tiny C Compiler passes more than one of them wrong. Errors: 6 for an invalid operation or an overflow,
   11 for a division by zero; VB6 stores the result (Inf, NaN), then raises (vb_fpstore). */
typedef long double vb_ld;
static inline vb_ld vb_fx(vb_ld v) { if (v != v || (v - v) != (v - v)) vb_fpraise(6); return v; }
static inline vb_ld vb_fxd(vb_ld v) { if (v != v) vb_fpraise(6); else if ((v - v) != (v - v)) vb_fpraise(11); return v; }
static inline double vb_rd(vb_ld v) {
    double r = (double)v;
    if (r == r && (r - r) != (r - r) && !((v - v) != (v - v))) vb_fpraise(6);   /* finite, too big for a Double */
    return r;
}
static inline float vb_rf(vb_ld v) {
    float r = (float)v;
    if (r == r && (r - r) != (r - r) && !((v - v) != (v - v))) vb_fpraise(6);
    return r;
}
float vb_chkf(double v);

double vb_round_even(double v);                /* banker's rounding */
vb_i16 vb_cint(double v);
vb_i32 vb_clng(double v);
vb_u8 vb_cbyte(double v);
vb_cy vb_ccur(double v);
float vb_csng(double v);
double vb_cy2d(vb_cy c);
vb_bool vb_cbool(double v);
double vb_fix(double v);
vb_i16 vb_abs_i(vb_i32 v);
vb_i32 vb_abs_l(vb_i64 v);
/* Remove Integer Overflow Checks (native code of VB6): two's complement, no error 6 */
static inline vb_i16 vb_abs_iw(vb_i32 v) { return (vb_i16)(v < 0 ? -v : v); }
static inline vb_i32 vb_abs_lw(vb_i64 v) { return (vb_i32)(vb_u32)(v < 0 ? -v : v); }
/* Byte + Byte into an Integer in native Fast code without On Error: no check either (VB6) */
static inline vb_u8 vb_add_bw(vb_i32 a, vb_i32 b) { return (vb_u8)(a + b); }
static inline vb_u8 vb_sub_bw(vb_i32 a, vb_i32 b) { return (vb_u8)(a - b); }
static inline vb_u8 vb_mul_bw(vb_i32 a, vb_i32 b) { return (vb_u8)(a * b); }
#ifndef __cplusplus
double fabs(double);
#endif
double vb_round(double v, vb_i32 digits);
double vb_int(double v);

vbs vb_cstr_i(vb_i32 v);
vbs vb_cstr_f(float v);
vbs vb_cstr_d(double v);
vbs vb_cstr_c(vb_cy v);
vbs vb_cstr_b(vb_bool v);
vbs vb_cstr_date(vb_date v);
vb_date vb_s2date(vbs s);
vbs vb_str_i(vb_i32 v);                        /* Str(): leading space */
vbs vb_str_f(float v);
vbs vb_str_d(double v);
vbs vb_str_c(vb_cy v);
double vb_val(vbs s);
double vb_s2d(vbs s);                          /* CDbl(string): error 13 if not a number */
vb_bool vb_s2bool(vbs s);

/* ---- objects -----------------------------------------------------------
 * An object starts with a vbobjh header; the fields of the class follow it.
 * Objects are reference counted like COM objects: Class_Terminate runs when the
 * last reference goes. Object expressions are borrowed references that live until
 * the end of the statement (like string temporaries); variables own a reference. */
typedef struct vb_class vb_class;
typedef struct vb_sink vb_sink;
/* A reference can also point to an interface "view" of an object (Implements): a view has
 * cls = the interface class and main = the object; VB6 keeps that identity, so late binding
 * through Object/Variant sees only the interface members (REA: t_events). */
typedef struct vbobjh {
    const vb_class *cls; vb_i32 refs; vb_i32 state; vb_sink *sinks;
    struct vbobjh *main;               /* view: the object itself; 0 for an object */
    struct vbobjh *views;              /* object: its views; view: the next view */
    void *ext;                         /* runtime data (forms, controls) */
} vbobjh;
typedef vbobjh *vbobj;
typedef void (*vb_fn)(void);           /* any procedure (tables of event handlers) */

/* ---- Variant ----------------------------------------------------------- */
enum { VT_EMPTY = 0, VT_NULL = 1, VT_I2 = 2, VT_I4 = 3, VT_R4 = 4, VT_R8 = 5, VT_CY = 6,
       VT_DATE = 7, VT_BSTR = 8, VT_DISPATCH = 9, VT_ERROR = 10, VT_BOOL = 11, VT_VARIANT = 12, VT_DECIMAL = 14, VT_UI1 = 17,
       VT_RECORD = 36, VT_ARRAY = 0x2000 };
struct vbrec;                          /* VT_RECORD: a Type in a Variant (u.rec, its own copy) */
struct vbarr;
/* vt = VT_ARRAY | element type: the Variant owns a copy of an array (u.a) */
typedef struct vbvar {
    vb_u16 vt; vb_u16 pad;             /* VT_DECIMAL: pad = scale | 0x8000 when negative */
    union { vb_i16 i; vb_i32 l; float f; double d; vb_cy cy; vb_u8 b; vbs s; vbobj o; struct vbarr *a; struct vbrec *rec;
            struct { vb_u32 hi, lo, mid; } dec; } u;   /* (12 bytes: the size of the union does not change) */
} vbvar;

/* ---- Decimal (decimal.c): a Variant only, as in VB6 --------------------- */
vbvar vb_cdec(vbvar v);                          /* CDec */
vbvar vb_dec_from_i64(vb_i64 x);
vbvar vb_dec_from_cy(vb_cy c);
vbvar vb_dec_from_double(double d, int digits);  /* 15 digits (Double), 7 (Single) */
vbvar vb_dec_from_str(vbs s);
vbvar vb_dec_from_ascii(const char *s);           /* "123.45" (a database): Empty if it does not fit */
double vb_dec2d(vbvar v);
vb_i64 vb_dec2i64(vbvar v, vb_i64 lo, vb_i64 hi); /* rounded half to even; error 6 outside lo..hi */
vb_cy vb_dec2cy(vbvar v);
vbs vb_dec2s(vbvar v);                           /* CStr */
void vb_dec_text(vbvar v, int locale, char *out);
vbvar vb_dec_add(vbvar a, vbvar b);
vbvar vb_dec_sub(vbvar a, vbvar b);
vbvar vb_dec_mul(vbvar a, vbvar b);
vbvar vb_dec_div(vbvar a, vbvar b);
int vb_dec_cmp(vbvar a, vbvar b);
vbvar vb_dec_neg(vbvar v);
vbvar vb_dec_abs(vbvar v);
int vb_dec_sgn(vbvar v);
vbvar vb_dec_int(vbvar v, int fix);
vbvar vb_dec_round(vbvar v, vb_i32 n);
void vb_dec_bytes(vbvar v, unsigned char *out);
vbvar vb_dec_from_bytes(const unsigned char *in);
vbvar vb_v_intfix(vbvar v, vb_i32 kind);         /* Int (0), Fix (1), Abs (2) of a Variant */
vb_i16 vb_v_sgn(vbvar v);
vbvar vb_v_round(vbvar v, vb_i32 n);
vbs vb_str_v(vbvar v);                           /* Str$ of a Variant */

/* late binding: every class has a table of its public members */
typedef vbvar (*vb_thunk)(vbobj self, vb_i32 argc, vbvar *argv);
enum { VBM_METHOD = 1, VBM_GET = 2, VBM_LET = 4, VBM_SET = 8 };
typedef struct vb_member { const char *name; vb_i32 kind; vb_thunk fn; } vb_member;  /* name: lower case, UTF-8 */
struct vb_class {
    const char *name;                  /* TypeName */
    vb_i32 size;
    void (*init)(vbobj);               /* field set-up + Class_Initialize */
    void (*term)(vbobj);               /* Class_Terminate */
    void (*clear)(vbobj);              /* frees the fields */
    const vb_member *members; vb_i32 nmembers;
    const char *defmember;             /* default member (lower case) or 0 */
    vb_i32 (*each_count)(vbobj);       /* For Each support (Collection) */
    vbvar (*each_item)(vbobj, vb_i32);
    const vb_class *const *ifaces;     /* Implements: 0-terminated list, or 0 */
    /* members the runtime provides (forms and controls); *found = 0 if there is none */
    vbvar (*fallback)(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
    const void *form;                  /* a form: its vbf_formdef */
    const char *enummember;            /* NewEnum of a class (VB_UserMemId = -4), lower case, or 0 */
    const char *progid;                /* an exposed class: "project.class", lower case; the same
                                          interface compiled into two units is one class */
};
vbobj vb_onew(const vb_class *c);              /* New c: a temporary reference */
void vb_oaddref(vbobj o);
void vb_orelease(vbobj o);
void vb_oset(vbobj *dst, vbobj src);           /* Set dst = src */
void vb_ofree(vbobj *v);                       /* release an object variable */
vbobj vb_otmp(vbobj o);                        /* an owned reference becomes a statement temporary */
vbobj vb_oauto(vbobj *v, const vb_class *c);   /* Dim v As New c: create on first use */
vbobj vb_ocast(vbobj o, const vb_class *c);
extern VB_IMPORT vbobj (*vb_obj_inner)(vbobj o);          /* the object behind a control (a user control) */    /* error 13 unless Nothing or a c */
void *vb_ofields(vbobj o, void *dummy);        /* error 91 if Nothing (returns dummy then) */
vb_bool vb_typeof(vbobj o, const vb_class *c); /* TypeOf o Is c (c = 0: Object) */
vbs vb_otypename(vbobj o);
vbvar vb_late(vbvar target, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv);  /* name "" = default member */
vbvar vb_late_o(vbobj target, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv);
vbvar vb_callbyname(vbobj target, vbs name, vb_i32 calltype, vb_i32 argc, vbvar *argv);   /* CallByName */
vbobj vb_v2obj(vbvar v);                       /* error 424 if v holds no object */
vbvar vb_v_obj(vbobj o);
vbvar vb_v_letval(vbvar v);                     /* v = w (Let): an object gives its default value */
vb_bool vb_typeof_name(vbobj o, const char *name);   /* TypeOf x Is TextBox / Control / Form */
vb_bool vb_isobject(vbvar v);
vb_bool vb_ismissing(vbvar v);
/* the Scripting runtime (scripting.c) and CreateObject */
extern VB_IMPORT const vb_class vb_cls_fso, vb_cls_drive, vb_cls_drives, vb_cls_folder, vb_cls_folders, vb_cls_file,
    vb_cls_files, vb_cls_textstream, vb_cls_dictionary;
/* DAO (runtime/db, libvbdb.a: linked when the project references DAO) */
extern VB_IMPORT const vb_class vb_cls_dao_dbengine, vb_cls_dao_workspace, vb_cls_dao_database, vb_cls_dao_tabledef,
    vb_cls_dao_field, vb_cls_dao_index, vb_cls_dao_relation, vb_cls_dao_querydef, vb_cls_dao_parameter,
    vb_cls_dao_recordset, vb_cls_dao_error, vb_cls_dao_coll, vb_cls_dao_property;
extern VB_IMPORT const vb_class vb_cls_ado_connection, vb_cls_ado_recordset, vb_cls_ado_fields, vb_cls_ado_field,
    vb_cls_ado_command, vb_cls_ado_parameters, vb_cls_ado_parameter, vb_cls_ado_errors, vb_cls_ado_error,
    vb_cls_ado_properties;
extern VB_IMPORT const vb_class vb_cls_bindingcollection, vb_cls_binding;
/* RDO 2.0 (runtime/db/rdo.inc): the objects of MSRDO20.DLL; rdoEngine, the one engine of the program */
extern VB_IMPORT const vb_class vb_cls_rdo_engine, vb_cls_rdo_environment, vb_cls_rdo_connection, vb_cls_rdo_query,
    vb_cls_rdo_resultset, vb_cls_rdo_column, vb_cls_rdo_parameter, vb_cls_rdo_table, vb_cls_rdo_error, vb_cls_rdo_coll;
vbobj vb_rdo_engine(void);
/* the Data Formatting Object Library (MSSTDFMT.DLL, stdfmt.c): StdDataFormat, StdDataValue, StdDataFormats, and
 * what a bound control does with its DataFormat: a value of the field into what it shows (the Format event) and
 * back (UnFormat: -1 when the text does not fit the format); vb_db_strto: a String into the type of an ADO
 * field with the regional settings (-1: it does not fit) */
extern VB_IMPORT const vb_class vb_cls_stddataformat, vb_cls_stddatavalue, vb_cls_stddataformats;
vbobj vb_dfmt_new(void);                                            /* a counted reference */
void vb_dfmt_load(vbobj f, vb_i32 type, vbs fmt, vb_i32 tfn, vbvar tv, vbvar fv, vbvar nv, vb_i32 day, vb_i32 week, vb_i32 lcid);
vb_bool vb_is_dfmt(vbobj o);
vb_i32 vb_dfmt_kind(vbobj o);                                        /* its Type; -1: none */
vbvar vb_dfmt_format(vbobj f, vbvar v, vbobj target);
vb_i32 vb_dfmt_unformat(vbobj f, vbvar v, vbobj target, vbvar *out);
vb_i32 vb_db_strto(vbvar v, vb_i32 adotype, vbvar *out);
vbobj vb_dao_dbengine(void);
/* the database runtime (libvbdb) as the other parts of the runtime see it, without linking it:
 * vb_db_init() of a program that uses a database fills these */
typedef struct vb_db_hooks {
    vbobj (*engine)(void);                                  /* DAO.DBEngine */
    /* the Data control is told before (what 1: Validate, *arg: 1 cancels) and after (what 2: Reposition)
     * its recordset moves; action: vbDataAction... */
    void (*rs_watch)(vbobj rs, void (*cb)(void *ctx, int what, int action, int *arg), void *ctx);   /* what 3: its rows
                                                               changed and it did not move (Delete) */
    vbobj (*createobject)(const char *progid);             /* "dao.dbengine.36" ... (lower case) */
    /* an ADO Recordset tells its listeners (the ADO Data control, bound controls, BindingCollection)
     * its events by name ("willmove" ...) and two of its own: "lb_save" before it saves or leaves the
     * current record (the bound values go into it), "lb_fill" when the current record is another or
     * its values changed (the bound controls show it); on = 0 stops */
    void (*ado_listen)(vbobj rs, void (*fn)(void *ctx, vbobj rs, const char *ev, int argc, vbvar *argv), void *ctx, int on);
    int (*is_ado_rs)(vbobj o);                             /* an ADO Recordset */
    /* PropertyChanged "Name" in a class (DataBindingBehavior): the BindingCollections that bind it */
    void (*prop_changed)(vbobj o, const char *name);
    /* the rows of an ADO Recordset for a grid or a list: rows (-1: closed), columns, the current row
     * (0-based; = rows while AddNew is on); a cell; a column; a move to a row by the user */
    int (*ado_shape)(vbobj rs, int *ncols, int *cur);
    vbvar (*ado_cell)(vbobj rs, int row, int col);
    const char *(*ado_colname)(vbobj rs, int col);
    int (*ado_coltype)(vbobj rs, int col);
    int (*ado_goto)(vbobj rs, int row);
    /* the rows of a DAO Recordset for the DBGrid, DBList and DBCombo (phase 7.6), as ado_shape ...: rows (the
     * records not deleted; -1: closed), columns, the current row (0-based; = rows while AddNew is on, -1 at
     * BOF / EOF); a cell (row -1: the record being added or edited); a column; the Bookmark of a row (Null:
     * none); the row of a Bookmark (-1: none) */
    int (*is_dao_rs)(vbobj o);
    int (*dao_shape)(vbobj rs, int *ncols, int *cur);
    vbvar (*dao_cell)(vbobj rs, int row, int col);
    const char *(*dao_colname)(vbobj rs, int col);
    int (*dao_coltype)(vbobj rs, int col);
    vbvar (*dao_bookmark)(vbobj rs, int row);
    int (*dao_bmrow)(vbobj rs, vbvar bm);
    /* the RemoteData control (phase 7.8): the ADO Recordset inside an rdoResultset (its rows: the bound controls
     * read and write them, and listen to it: "lb_validate", "lb_fill", "lb_reposition"), or 0 */
    vbobj (*rdo_inner)(vbobj rdors);
    /* the DataRepeater (phase 7.5): the Bookmark of a row of an ADO Recordset as ado_shape counts them (Null: none);
     * the row of a Bookmark (-1: none). An ADO Recordset also tells its listeners "lb_willchange" (argv[0]: the
     * column) before a field takes a new value */
    vbvar (*ado_bookmark)(vbobj rs, int row);
    int (*ado_bmrow)(vbobj rs, vbvar bm);
    /* a control is bound to the Recordset through a data source object (a class: Set Text1.DataSource = obj): the
     * first such binding fetches the row, and the Recordset tells MoveComplete (VB6 oracle DEV-9 df3-15) */
    void (*ado_consumer)(vbobj rs);
} vb_db_hooks;
extern VB_IMPORT vb_db_hooks vb_dbhooks;
void vb_prop_changed(vbobj o, vbs name);               /* PropertyChanged "Name" in a class */
void vb_db_init(void);
vbobj vb_createobject(vbs progid, vbvar server);
/* components (ActiveX DLL / EXE projects compiled into the program, component.c) */
typedef struct vb_progclass { const char *name; const vb_class *cls; vb_i32 flags; } vb_progclass;   /* name: lower case */
enum { VBC_CREATABLE = 1, VBC_GLOBAL = 2, VBC_SINGLE = 4 };
/* an ActiveX EXE with SingleUse classes: each of their objects has a copy of the server of its own (its
 * module variables, its Sub Main); the calls into an object run with the globals of its copy */
typedef struct vb_ctxdef {
    unsigned long (*size)(void);        /* the bytes of the globals */
    void (*swap)(void *buf);            /* the globals <-> buf */
    void (*init)(void);                 /* the initialisations of the globals (fixed arrays ...) */
    void (*main)(void);                 /* Sub Main of the server, or 0 */
    void *cur;                          /* the copy whose globals are in place; 0: the first copy */
    int depth;                          /* > 0: the server's own code runs (its New is no client's) */
    void *priv;                         /* the runtime's list of the copies */
} vb_ctxdef;
void *vb_ctx_begin(vbobj o, vb_ctxdef *d, int single);   /* Class_Initialize: the copy of a new object */
void *vb_ctx_enter(vbobj o, vb_ctxdef *d);              /* a call into o */
void vb_ctx_end(vb_ctxdef *d, void *prev);
void vb_register_component(const char *proj, const vb_progclass *classes, vb_i32 n, void (*init)(void));
vbobj vb_component_create(const char *lowprogid, int *found);
void vb_err_source_default(const char *project);   /* an error leaving a component: Source = its name */
vbobj vb_component_new(const char *progid);          /* New Proj.Class */
vbobj vb_component_global(const char *progid);       /* a GlobalMultiUse class: its one instance */
const vb_class *vb_component_class(const char *progid);  /* Dim x As New Proj.Class */
/* AddressOf: the number of a procedure; the callers of callbacks (SetTimer ...) call it through vb_callback */
typedef vb_i32 (*vb_cbfn)(vb_i32 argc, vb_i32 *argv);
vb_i32 vb_addressof(vb_cbfn f);
vb_cbfn vb_callback_fn(vb_i64 h);                    /* 0 if h is no AddressOf number */
vb_i32 vb_callback(vb_i64 h, vb_i32 argc, vb_i32 *argv);
void vb_each_begin(vbvar *v);                  /* For Each: an object with NewEnum gives its enumerator */
vb_i32 vb_each_count(vbvar v);                 /* For Each over a Collection */
vbvar vb_each_item(vbvar v, vb_i32 i);
extern VB_IMPORT const vb_class vb_cls_collection;
/* a class that is a data source (DataSourceBehavior = vbDataSource): DataMembers (a Collection of the names it
 * gives) and DataMemberChanged "name" (the bindings read the member again: BindingCollection, bind.c) */
vbvar vb_ds_datamembers(vbobj self, vb_i32 argc, vbvar *argv);
vbvar vb_ds_datamemberchanged(vbobj self, vb_i32 argc, vbvar *argv);
extern VB_IMPORT void (*vb_ds_changed_hook)(vbobj source, const char *member);
extern VB_IMPORT void (*vb_ds_changed_ctls_hook)(vbobj source, const char *member);   /* the forms runtime */
vb_bool vb_implements(const vb_class *cls, const vb_class *iface);
vbobj vb_oid(vbobj o);                         /* the object behind a reference (Is, method calls) */
vbobj vb_oview(vbobj o, const vb_class *iface); /* the reference for an implemented interface */

/* events: a WithEvents variable connects its object (the source) to the object that holds
 * the variable (the listener); the source keeps a weak list of listeners. */
void vb_wev_set(vbobj listener, vbobj *field, vbobj src, const vb_fn *handlers);   /* Set field = src */
vb_i32 vb_sinks_begin(vbobj src, vbobj **listeners, const vb_fn ***tables);       /* RaiseEvent: snapshot */
void vb_sinks_end(vbobj *listeners, const vb_fn **tables, vb_i32 n);

vbvar vb_v_empty(void);
vbvar vb_v_null(void);
vbvar vb_v_missing(void);
vbvar vb_cverr(vb_i32 n);                       /* CVErr: a Variant of the type Error */
vbvar vb_v_i(vb_i16 v);
vbvar vb_v_l(vb_i32 v);
vbvar vb_v_f(float v);
vbvar vb_v_d(double v);
vbvar vb_v_c(vb_cy v);
vbvar vb_v_date(vb_date v);
vbvar vb_v_bool(vb_bool v);
vbvar vb_v_byte(vb_u8 v);
vbvar vb_v_s(vbs v);                           /* refers to v (temporary) */
void vb_vset(vbvar *dst, vbvar src);           /* variable = value (copies strings) */
void vb_vfree(vbvar *v);
vbvar vb_vtmp(vbvar v);                        /* temporary copy (strings too) */

vb_i16 vb_v2i(vbvar v);
vb_i32 vb_v2l(vbvar v);
float vb_v2f(vbvar v);
double vb_v2d(vbvar v);
vb_cy vb_v2c(vbvar v);
vb_bool vb_v2bool(vbvar v);
vb_u8 vb_v2byte(vbvar v);
vbs vb_v2s(vbvar v);
vb_date vb_v2date(vbvar v);

vbvar vb_v_add(vbvar a, vbvar b);
vbvar vb_v_sub(vbvar a, vbvar b);
vbvar vb_v_mul(vbvar a, vbvar b);
vbvar vb_v_div(vbvar a, vbvar b);
vbvar vb_v_idiv(vbvar a, vbvar b);
vbvar vb_v_mod(vbvar a, vbvar b);
vbvar vb_v_pow(vbvar a, vbvar b);
vbvar vb_v_neg(vbvar a);
vbvar vb_v_cat(vbvar a, vbvar b);
vbvar vb_v_and(vbvar a, vbvar b);
vbvar vb_v_or(vbvar a, vbvar b);
vbvar vb_v_xor(vbvar a, vbvar b);
vbvar vb_v_eqv(vbvar a, vbvar b);
vbvar vb_v_imp(vbvar a, vbvar b);
vbvar vb_v_not(vbvar a);
vbvar vb_v_cmp(vbvar a, vbvar b, vb_i32 op, vb_i32 flags); /* op: 0 = 1 <> 2 < 3 > 4 <= 5 >=; flags 1 text, 2/4 typed side */
vb_bool vb_v_truth(vbvar v);                   /* If <variant> */
vb_i32 vb_vartype(vbvar v);
vb_i32 vb_vartype_def(vbvar v);                 /* IsNull, IsEmpty: of the default value of an object */
vbs vb_typename(vbvar v);
vb_bool vb_isnumeric(vbvar v);
vb_bool vb_isdate(vbvar v);
vbs vb_error_str(vb_i32 n);                     /* Error(n) */

/* ---- arrays ------------------------------------------------------------ */
enum { VBA_PLAIN = 0, VBA_STR = 1, VBA_VAR = 2, VBA_OBJ = 3, VBA_UDT = 4 };
/* a user-defined type (Type ... End Type): how to set up, free and copy one */
/* kinds of data for Get/Put (and the fields of a Type) */
enum { VBK_I2 = 1, VBK_I4, VBK_R4, VBK_R8, VBK_CY, VBK_DATE, VBK_BOOL, VBK_U8, VBK_STR, VBK_FSTR, VBK_VAR, VBK_UDT, VBK_OBJ };
typedef struct vb_udt vb_udt;
void vb_cb_udt_in(const vb_udt *u, void *dst, vb_i32 addr);    /* AddressOf: a ByRef Type parameter */
void vb_cb_udt_out(const vb_udt *u, const void *src, vb_i32 addr);
vbobj vb_cb_obj(vb_i32 addr);
/* arr: 0 a value, 1 a fixed array, 2 a dynamic array */
typedef struct vb_ufield { vb_i32 kind, fixlen, offset, arr; const vb_udt *udt; } vb_ufield;
struct vb_udt {
    vb_i32 size; void (*init)(void *); void (*free)(void *); void (*copy)(void *, const void *);
    const vb_ufield *fields; vb_i32 nfields;
    const char *name;                  /* "Project.Type" of a Public Type of a public class (it can go in a Variant);
                                        * the same name in a component and in the program that uses it: the same Type */
    const char *const *fnames;         /* (such a Type) the names of its fields: v.Qty of a Variant holding one */
};
/* a Type in a Variant (VB6: only a Public Type of a public class or user control of an ActiveX project) */
vbvar vb_v_rec(const void *p, const vb_udt *u);       /* a Variant holding a copy (a temporary, as vb_v_s) */
void vb_v2rec(vbvar v, const vb_udt *u, void *dst);
vbvar vb_rec_member(vbvar v, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv);   /* v.Field (late) */   /* the Type in the Variant into dst (error 13: not that Type) */
struct vbrec *vb_rec_new(const vb_udt *u, const void *src);   /* a copy of src (0: an empty one) */
struct vbrec *vb_rec_tmp(const vb_udt *u, const void *src);
int vb_rec_istmp(const struct vbrec *r);              /* (vb_vfree leaves those to vb_release, as strings) */   /* the same, freed at the end of the statement */
void vb_rec_free(struct vbrec *r);
const vb_udt *vb_rec_udt(const struct vbrec *r);
void *vb_rec_data(struct vbrec *r);
#define VB_OFFSETOF(T, f) ((vb_i32)((char *)&((T *)0)->f - (char *)0))
typedef struct vbarr {
    char *data;
    vb_i32 nd;        /* dimensions, 0 = not dimensioned (dynamic) */
    vb_i32 esz;       /* element size                              */
    vb_i32 kind;      /* VBA_PLAIN, VBA_STR, VBA_VAR, VBA_OBJ, VBA_UDT */
    vb_i32 fixed;     /* 1 = Dim a(10): cannot be ReDim'ed          */
    vb_i32 n;
    vb_i32 lb[8], ub[8], mul[8];
    const vb_udt *udt;                 /* VBA_UDT: the element type (set before vb_adim) */
} vbarr;

void vb_adim(vbarr *a, vb_i32 esz, vb_i32 kind, vb_i32 fixed, vb_i32 nd, ...);   /* lb, ub pairs */
void vb_aredim(vbarr *a, vb_i32 esz, vb_i32 kind, vb_i32 preserve, vb_i32 nd, ...);
void vb_aerase(vbarr *a);
void vb_afree(vbarr *a);
void vb_acopy(vbarr *dst, const vbarr *src);     /* deep copy (arrays inside a Type) */
void *vb_aelem_lin(vbarr *a, vb_i32 i);           /* i-th element in memory order (For Each) */
vbarr *vb_atmp(void);                             /* a new empty array, freed at the end of the statement */
void vb_aempty(vbarr *a, vb_i32 esz, vb_i32 kind); /* a dimensioned array with no elements: LBound 0, UBound -1 */
void vb_amove(vbarr *dst, vbarr *src);            /* dst = src, src becomes empty (function results) */
void vb_aparam(vbarr *dst, vb_i32 n, vbvar *items); /* ParamArray: Variant array 0 .. n-1 */
/* arrays inside a Variant (vt = VT_ARRAY | element VarType) */
vbvar vb_v_arr(const vbarr *a, vb_i32 elemvt);    /* a temporary Variant holding a copy */
vbarr *vb_v_arrptr(vbvar v);                      /* the array of a Variant (error 13 if none) */
void vb_v_erase(vbvar *v);                        /* Erase v (a Variant): the array erased, kept */
void vb_v2arr(vbvar v, vbarr *dst, vb_i32 elemvt);
void vb_afix_init(vbarr *a, vb_i32 n);           /* the elements of an array of String * n */
void vb_s2barr(vbs s, vbarr *dst);             /* Byte array = String */
vbs vb_barr2s(const vbarr *a);                 /* String = Byte array *//* typed array = Variant */
vbvar vb_v_index(vbvar v, vb_i32 argc, vbvar *argv);           /* v(i, ...) */
void vb_v_setindex(vbvar *v, vbvar value, vb_i32 argc, vbvar *argv, vb_i32 set);
vb_bool vb_isarray(vbvar v);
vbarr vb_split(vbs s, vbs delim, vb_i32 has_delim, vb_i32 limit, vb_i32 text);
vbs vb_join(vbarr *a, vb_i32 elemvt, vbs delim, vb_i32 has_delim);
vbarr vb_filter(vbarr *a, vb_i32 elemvt, vbs match, vb_i32 include, vb_i32 text);
vbvar vb_array(vb_i32 n, vbvar *items);           /* Array(...) */
void *vb_aelem(vbarr *a, vb_i32 nd, ...);
vb_i32 vb_lbound(vbarr *a, vb_i32 dim);
vb_i32 vb_ubound(vbarr *a, vb_i32 dim);
static inline void *vb_aelem1(vbarr *a, vb_i32 i) {
    extern void *vb_aelem_err(vbarr *a);
    if (a->nd != 1 || i < a->lb[0] || i > a->ub[0]) return vb_aelem_err(a);
    return a->data + (vb_i64)(i - a->lb[0]) * a->esz;
}
void *vb_aelem_nc(vbarr *a, vb_i32 nd, ...);      /* Remove Array Bounds Checks */
static inline void *vb_aelem1_nc(vbarr *a, vb_i32 i) {
    if (a->nd != 1 || i < a->lb[0] || i > a->ub[0]) return vb_aelem_nc(a, 1, i);
    return a->data + (vb_i64)(i - a->lb[0]) * a->esz;
}

/* ---- output: Debug.Print, Print #, Write # ---------------------------- */
void vb_pr_begin(vb_i32 fileno);               /* 0 = Debug (standard output) */
void vb_pr_s(vbs s);
void vb_pr_i(vb_i32 v);
void vb_pr_f(float v);
void vb_pr_d(double v);
void vb_pr_c(vb_cy v);
void vb_pr_bool(vb_bool v);
void vb_pr_date(vb_date v);
void vb_pr_v(vbvar v);
void vb_pr_comma(void);
void vb_pr_tab(vb_i32 col);                    /* col <= 0: next print zone */
void vb_pr_spc(vb_i32 n);
void vb_pr_end(vb_i32 newline);
void vb_wr_begin(vb_i32 fileno);
void vb_wr_v(vbvar v);
void vb_wr_end(void);

/* ---- files ------------------------------------------------------------- */
enum { VBF_INPUT = 1, VBF_OUTPUT = 2, VBF_APPEND = 8, VBF_BINARY = 32, VBF_RANDOM = 4 };
void vb_open(vbs path, vb_i32 mode, vb_i32 fileno, vb_i32 reclen);
void vb_close(vb_i32 fileno);                  /* 0 = all */
/* Get / Put (Binary and Random files); rec <= 0: the current position */
void vb_put(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, const void *p);
void vb_get(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, void *p);
void vb_put_arr(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, vbarr *a, vb_i32 dynamic);
void vb_get_arr(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, vbarr *a, vb_i32 dynamic);
void vb_seek(vb_i32 fileno, vb_i32 pos);       /* Seek #f, pos */
vb_i32 vb_seekpos(vb_i32 fileno);              /* Seek(f) */
vb_i32 vb_loc(vb_i32 fileno);
void vb_line_input(vb_i32 fileno, vbs *dst);
void vb_input_v(vb_i32 fileno, vbvar *dst);    /* one item of Input # */
void vb_input_s(vb_i32 fileno, vbs *dst);
vbs vb_input_n(vb_i32 n, vb_i32 fileno);        /* Input(n, #f) */
vbs vb_inputb(vb_i32 n, vb_i32 fileno);         /* InputB(n, #f) */
void vb_input_d(vb_i32 fileno, double *dst);
vb_bool vb_eof(vb_i32 fileno);
vb_i32 vb_lof(vb_i32 fileno);
vb_i32 vb_freefile(void);
void vb_kill(vbs path);
vb_i32 vb_filelen(vbs path);
void vb_mkdir(vbs path);
void vb_rmdir(vbs path);
void vb_chdir(vbs path);
void vb_alock(const vbarr *a);                 /* an element passed ByRef: ReDim / Erase give 10 */
void vb_aunlock(const vbarr *a);
int vb_alocked(const vbarr *a);
vbs vb_curdir(void);
vbs vb_curdir_d(vbs drive);                    /* CurDir(drive) */
vb_i32 vb_recno(vb_i32 rec);                   /* a record number of Put/Get (0: error 63) */
vbs vb_dir(vbs pattern, vb_i32 attr, vb_i32 has_pattern);
void vb_name(vbs from, vbs to);
vbs vb_app_path(void);
vbs vb_app_exename(void);
vbobj vb_err_obj(void);                       /* Err as an object */
vbobj vb_global_obj(void);                    /* VB.Global as an object */
vbobj vb_app_obj(void);                       /* App as an object */
void vb_app_info(const char *key, vbs v);     /* what the .vbp says of the program (vb_app_obj) */
vbs vb_app_str(int which, vbs dflt);                /* App.Title (1), App.HelpFile (2): as set at run time, or dflt */
vbs vb_app_helpfile(void);                              /* App.HelpFile now (the .vbp's, or as set at run time) */
void vb_app_set_str(int which, vbs v);
vbs vb_command(void);
void vb_filecopy(vbs from, vbs to);
vb_date vb_filedatetime(vbs path);
vb_i32 vb_getattr(vbs path);
void vb_setattr(vbs path, vb_i32 attr);
void vb_chdrive(vbs drive);
void vb_appactivate(vbs title, vb_i32 wait);
extern VB_IMPORT void (*vb_appactivate_hook)(vbs title);
double vb_pmt(double rate, double nper, double pv, double fv, vb_i32 type);
double vb_fv(double rate, double nper, double pmt, double pv, vb_i32 type);
double vb_pv(double rate, double nper, double pmt, double fv, vb_i32 type);
double vb_ipmt(double rate, double per, double nper, double pv, double fv, vb_i32 type);
double vb_ppmt(double rate, double per, double nper, double pv, double fv, vb_i32 type);
double vb_nper(double rate, double pmt, double pv, double fv, vb_i32 type);
double vb_sln(double cost, double salvage, double life);
double vb_syd(double cost, double salvage, double life, double period);
double vb_ddb(double cost, double salvage, double life, double period, double factor);
double vb_npv(double rate, const vbarr *a, vb_i32 vt);
double vb_irr(const vbarr *a, vb_i32 vt, double guess);
double vb_mirr(const vbarr *a, vb_i32 vt, double finance, double reinvest);
double vb_rate(double nper, double pmt, double pv, double fv, vb_i32 due, double guess);
vbs vb_partition(vb_i32 number, vb_i32 start, vb_i32 stop, vb_i32 interval);
vb_i32 vb_fileattr(vb_i32 fileno, vb_i32 type);
double vb_shell(vbs cmd, vb_i32 style);
vbs vb_environ_s(vbs name);
vbs vb_environ_n(vb_i32 n);

/* ---- VBA library ------------------------------------------------------- */
vb_i32 vb_len(vbs s);
vbs vb_left(vbs s, vb_i32 n);
vbs vb_right(vbs s, vb_i32 n);
vbs vb_mid(vbs s, vb_i32 start, vb_i32 n);     /* n < 0: rest of the string */
void vb_mid_stmt(vbs *dst, vb_i32 start, vb_i32 n, vbs src);
vb_i32 vb_instr(vb_i32 start, vbs a, vbs b, vb_i32 text);
vb_i32 vb_instrrev(vbs a, vbs b, vb_i32 start, vb_i32 text);
vbs vb_lcase(vbs s);
vbs vb_ucase(vbs s);
vbs vb_ltrim(vbs s);
vbs vb_rtrim(vbs s);
vbs vb_trim(vbs s);
vbs vb_space(vb_i32 n);
vbs vb_string_n(vb_i32 n, vb_i32 ch);
vbs vb_string_s(vb_i32 n, vbs ch);
vbs vb_chr(vb_i32 code);                       /* ANSI code page */
vbs vb_chrw(vb_i32 code);
vb_i32 vb_asc(vbs s);
vb_i32 vb_ascw(vbs s);
vbs vb_strreverse(vbs s);
vbs vb_replace(vbs s, vbs find, vbs with, vb_i32 start, vb_i32 count, vb_i32 text);
vbs vb_hex(vb_i64 v, vb_i32 bytes);
vbs vb_oct(vb_i64 v, vb_i32 bytes);
vb_i32 vb_strcomp(vbs a, vbs b, vb_i32 text);
vbs vb_lset(vbs dst, vbs src);
double vb_rnd(double n, vb_i32 has_n);
void vb_randomize(double seed, vb_i32 has_seed);
double vb_timer(void);
double vb_sgn(double v);
double vb_sqr(double v);
double vb_log(double v);
double vb_exp(double v);
double vb_sin(double v);
double vb_cos(double v);
double vb_tan(double v);
double vb_atn(double v);
vb_date vb_now(void);
vb_date vb_dateserial(vb_i32 y, vb_i32 m, vb_i32 d);
vb_date vb_timeserial(vb_i32 h, vb_i32 m, vb_i32 s);
vb_i32 vb_year(vb_date d);
vb_i32 vb_month(vb_date d);
vb_i32 vb_day(vb_date d);
vb_i32 vb_hour(vb_date d);
vb_i32 vb_minute(vb_date d);
vb_i32 vb_second(vb_date d);
vb_i32 vb_weekday(vb_date d, vb_i32 firstday);
vb_bool vb_like(vbs s, vbs pattern, vb_i32 text);
vb_date vb_dateadd(vbs interval, double n, vb_date d);
vb_i32 vb_datediff(vbs interval, vb_date a, vb_date b, vb_i32 firstday);
vb_i32 vb_datepart(vbs interval, vb_date d, vb_i32 firstday, vb_i32 firstweek);
vb_date vb_datevalue(vbs s);
vb_date vb_timevalue(vbs s);
/* byte strings: a String with any number of bytes (VBS_ODD) */
vb_i32 vb_sbytes(vbs s);                       /* LenB */
int vb_sbyte(vbs s, vb_i32 i);                 /* byte i, from 0 */
vbs vb_sfrombytes(const char *b, vb_i32 n);
vbs vb_sfrombytes2(vbs s, vb_i32 from, vb_i32 n);
vb_i32 vb_lenb(vbs s);
vbs vb_chrb(vb_i32 code);
vb_i32 vb_ascb(vbs s);
vbs vb_leftb(vbs s, vb_i32 n);
vbs vb_rightb(vbs s, vb_i32 n);
vbs vb_midb(vbs s, vb_i32 start, vb_i32 n);
vb_i32 vb_instrb(vb_i32 start, vbs s, vbs find, vb_i32 text);
void vb_midb_stmt(vbs *s, vb_i32 start, vb_i32 n, vbs value);
vbs vb_strconv(vbs s, vb_i32 conv);
vbs vb_strconv_l(vbs s, vb_i32 conv, vb_i32 lcid, vb_i32 has_lcid);   /* StrConv(s, conv, LocaleID) */
/* format.c */
vbs vb_format(vbvar v, vbs fmt, vb_i32 has_fmt, vb_i32 firstday, vb_i32 firstweek);
vbs vb_formatnumber(vbvar v, vb_i32 digits, vb_i32 lead0, vb_i32 parens, vb_i32 group, vb_i32 kind); /* kind 0 number 1 currency 2 percent */
vbs vb_formatdatetime(vb_date d, vb_i32 named);
vbs vb_monthname(vb_i32 m, vb_i32 abbrev);
vbs vb_weekdayname(vb_i32 w, vb_i32 abbrev, vb_i32 firstday);
vb_i32 vb_msgbox(vbs prompt, vb_i32 buttons, vbs title);
vb_i32 vb_msgbox_help(vbs prompt, vb_i32 buttons, vbs title, vbs helpfile, vb_i32 context);   /* MsgBox with HelpFile, Context */
vb_i32 vb_msgbox_helpinfo(vbs *helpfile);              /* the forms runtime: the help of the MsgBox shown (its Help button) */
vbs vb_inputbox(vbs prompt, vbs title, vbs def);
vbs vb_getsetting(vbs app, vbs section, vbs key, vbs def);
void vb_savesetting(vbs app, vbs section, vbs key, vbs value);
void vb_deletesetting(vbs app, vbs section, vbs key, vb_i32 has_key);
vbvar vb_getallsettings(vbs app, vbs section);
/* the resource file of the project: type and name are a number, or a name (upper case) when not 0 */
typedef struct vb_res { vb_i32 type_id; const char *type_name; vb_i32 id; const char *name;
                        const unsigned char *data; vb_i32 size; } vb_res;
void vb_res_init(const vb_res *table, vb_i32 n);
vbs vb_loadresstring(vb_i32 id);
vbvar vb_loadresdata(vbvar id, vbvar type);
vbobj vb_loadrespicture(vbvar id, vb_i32 format);
extern VB_IMPORT vbobj (*vb_picture_data_hook)(const char *data, vb_i32 n, vb_i32 type);
void vb_beep(void);
void vb_doevents(void);

/* ---- Declare (DLL procedures) ------------------------------------------
 * Windows APIs that programs use often are provided by the runtime (kernel32 ...);
 * other libraries are loaded as shared objects. Addresses from VarPtr/StrPtr/ObjPtr are
 * 32-bit handles into a map of real addresses, so they fit in a Long like in VB6. */
enum { VBD_INT = 1, VBD_DBL, VBD_PTR, VBD_STR, VBD_UDT,     /* a value, a double, ByRef pointer, ByVal String,
                                                               a Type (p: the value, s: its vb_udt) */
       VBD_SREF, VBD_VAR, VBD_BADCC, VBD_OREF };   /* VBD_OREF: an object variable ByRef to As Any (p: vbobj *) */                      /* ByRef String (s: the variable), ByRef Variant (p),
                                                               a Variant ByVal to As Any (error 49) */
typedef struct vb_dllarg { vb_i32 kind; vb_i64 i; double d; void *p; vbs *s; } vb_dllarg;
typedef struct vb_declare { const char *lib, *entry; void *fn; vb_i32 state; } vb_declare;
typedef vb_i64 (*vb_apiext)(vb_i32 argc, vb_dllarg *argv);
extern VB_IMPORT vb_apiext (*vb_api_hook)(const char *lib, const char *name);   /* more Windows procedures (user32 of the forms runtime) */
void *vb_dllarg_ptr(vb_dllarg *a, vb_i32 size);           /* the address an argument stands for */
char *vb_dllarg_str(vb_dllarg *a);
void vb_udt_lset(void *d, const vb_udt *du, const void *s, const vb_udt *su);   /* LSet a = b (Types) */
void vb_slset(vbs *dst, vbs src, vb_i32 right);                              /* LSet / RSet s = v */                         /* the ANSI text of a String argument (or address) */
vb_i64 vb_dllcall(vb_declare *d, vb_i32 argc, vb_dllarg *argv);
double vb_dllcall_d(vb_declare *d, vb_i32 argc, vb_dllarg *argv);
vbs vb_dllcall_s(vb_declare *d, vb_i32 argc, vb_dllarg *argv);   /* As String: the BSTR (ANSI) as a String */
extern VB_IMPORT vb_i32 vb_lastdllerr;                     /* Err.LastDllError */
vb_i32 vb_varptr_elem(const void *p, vb_i32 size, const vbarr *a);   /* VarPtr(a(0)): the rest too */
vb_i32 vb_varptr(const void *p, vb_i32 size);
vb_i32 vb_varptr_str(const vbs *p);           /* VarPtr of a String variable (its 4-byte BSTR pointer for CopyMemory) */
vb_i32 vb_strptr(vbs s);
void *vb_ptr_resolve(vb_i64 addr, vb_i32 size);  /* real address of a VarPtr/StrPtr value (0: error 5) */

/* ---- MCI (mci.c): the multimedia commands of winmm; Linux has its own MCI, Windows the real one */
typedef void (*vb_mci_notify)(void *ctx, vb_i32 devid, vb_i32 code);   /* MM_MCINOTIFY: 1 ok, 2 superseded, 4 aborted, 8 failure */
typedef struct { vb_i64 hwnd; int show, playing, fresh, w, h, stride; const unsigned char *data; char title[256]; } vb_mci_frameinfo;
vb_i32 vb_mci_send(const char *cmd, char *ret, vb_i32 retlen, vb_mci_notify cb, void *ctx);   /* UTF-8; 0 or the MCI error */
vb_i32 vb_mci_send_hwnd(const char *cmd, char *ret, vb_i32 retlen, vb_mci_notify cb, void *ctx, vb_i64 hwnd);
vb_i32 vb_mci_error_text(vb_i32 code, char *buf, vb_i32 len);   /* 1 if the code is an MCI error */
vb_i32 vb_mci_device_id(const char *name);
void vb_mci_poll(void);                 /* deliver the notifications that are due */
void vb_mci_forget(void *ctx);          /* a control is gone: no more notifications for it */
vb_i32 vb_mci_active(void);             /* devices playing or waiting to notify */
vb_i32 vb_mci_next_video(vb_i32 after); /* devices with pictures: 0 starts, 0 ends */
int vb_mci_frame(vb_i32 id, vb_mci_frameinfo *f);
extern VB_IMPORT void (*vb_mci_post_hook)(vb_i64 hwnd, vb_i32 code, vb_i32 devid);   /* mciSendString notify to a window */

/* ---- forms (libvbforms, Qt6) -------------------------------------------
 * A form module is a class whose objects also have the built-in members of a form
 * (Caption, Show, controls by name ...). Design-time properties come from the .frm:
 * each value is "s<text>" (a string), "n<number>" or "r<raw text>". */
typedef struct vbf_ctldef { const char *cls; const char *name; vb_i32 index; vb_i32 parent; const char *const *props; } vbf_ctldef;
typedef struct vbf_formdef {
    const char *name; const char *const *props;
    vb_i32 nctls; const vbf_ctldef *ctls;
    const vb_member *events; vb_i32 nevents;   /* handlers: "command1_click", "form_load" ... */
    vb_i32 mdi;                        /* 1: an MDIForm */
} vbf_formdef;
vbvar vbf_form_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
void vbf_form_clear(vbobj o);
vbobj vb_odefault(vbobj *slot, const vb_class *cls);    /* the default instance of a predeclared class */
vbobj vbf_default(vbobj *slot, const vb_class *cls);
vbobj vbf_report_new(vbobj owner, vbs def);            /* a DataReport (forms/report.inc) */   /* the default instance (Form1) */
vbobj vbf_ctl(vbobj form, const char *name);           /* a control or control array of a form */
void vbf_load(vbobj o);                                  /* Load Form1 / Load Text1(3) */
void vbf_load_elem(vbobj arr, vb_i32 index);
void vbf_unload(vbobj o);
vbobj vbf_forms(void);
void vbf_mdi_register(vbobj *slot, const vb_class *cls); /* the MDIForm of the project (its default instance) */
void vbf_uctl_register(const char *progid, const vb_class *cls);   /* a user control: "project.name" */
void vbf_uctl_event(vbobj uc, const char *event, vb_i32 argc, vbvar *argv);   /* RaiseEvent: to its form */
void vb_fire_named(vbobj src, const char *event, vb_i32 argc, vbvar *argv);    /* RaiseEvent of a form (object.c) */
vbvar vbf_uc_member(vbobj uc, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv);   /* UserControl.X */
vbobj vbf_screen(void);                                  /* the Screen object */
vbobj vbf_clipboard(void);                               /* the Clipboard object */
vbobj vbf_licenses(void);                                /* the Licenses collection */
vbobj vbf_printer(void);                                 /* the Printer object (a PDF file) */
vbobj vbf_printers(void);
void vbf_start(vbobj form);                              /* Startup = a form */
void vbf_run(void);                                      /* runs until no form is loaded */
void vbf_use(void);                                      /* links the forms runtime */
const vb_class *vbf_cls_stdfont(void);                   /* New StdFont */
const vb_class *vbf_cls_stdpicture(void);                /* New StdPicture */
const vb_class *vbf_cls_propertybag(void);               /* New PropertyBag */
extern VB_IMPORT void (*vb_doevents_hook)(void);
extern VB_IMPORT void (*vb_sendkeys_hook)(vbs keys, vb_i32 wait);
void vb_sendkeys(vbs keys, vb_bool wait);
extern VB_IMPORT vbobj (*vb_loadpicture_hook)(vbs path, vb_i32 has_path);
extern VB_IMPORT void (*vb_savepicture_hook)(vbobj pic, vbs path);
vbobj vb_loadpicture(vbs path, vb_i32 has_path);     /* a StdPicture object */
void vb_savepicture(vbobj pic, vbs path);
extern VB_IMPORT vb_i32 (*vb_msgbox_hook)(vbs prompt, vb_i32 buttons, vbs title);      /* the forms runtime: dialogs */
extern VB_IMPORT vbs (*vb_inputbox_hook)(vbs prompt, vbs title, vbs def);
vb_i32 vb_rgb(vb_i32 r, vb_i32 g, vb_i32 b);
vb_i32 vb_qbcolor(vb_i32 c);
vbs vb_pr_take(vb_i32 *newline);               /* the text of a Print method (vb_pr_begin(-2)) */

/* ---- program ----------------------------------------------------------- */
void vb_init(int argc, char **argv);

/* ---- the debugger (lbvc --debug, debug.c) ----------------------------- */
typedef struct vb_dbgvar { const char *name; vb_i32 vt; void *p; } vb_dbgvar;  /* vt: VT_*, VT_ARRAY|elem (p: vbarr), 36 a Type */
typedef struct vb_dbgframe {
    const char *proc; vb_i32 mod, line, nvars; const vb_dbgvar *vars;
    vb_i32 nmvars; const vb_dbgvar *mvars;                      /* the variables of the module (Me)   */
    struct vb_dbgframe *parent; vb_i32 depth; vb_i32 *ehmode, *inh; vbobj self;
    const vb_i32 *lines; vb_i32 nlines;        /* (line, statement) pairs: Set Next Statement             */
    vb_i32 *cur; vb_i32 jump;                  /* the statement now; the one to go to (0: none)           */
} vb_dbgframe;
extern VB_IMPORT int vb_dbg_on;
void vb_dbg_init(const char *const *mods, int nmods, const vb_dbgvar *globals, int nglobals);
void vb_dbg_enter(vb_dbgframe *f);
void vb_dbg_leave(vb_dbgframe *f);
void vb_dbg_line(vb_dbgframe *f, vb_i32 line);
void vb_dbg_stop(vb_dbgframe *f);
void vb_dbg_error(vb_dbgframe *f);
void vb_dbg_trap(vb_dbgframe *f);              /* an error the procedure handles: Break on All Errors     */
void vb_dbg_kinds(const char *kinds);          /* a letter a module: C class, U user control, F form, M module */
vb_i32 vb_err_number_pending(void);
void vb_err_clear_pending(void);
void vb_set_project(const char *title);         /* Err.Source of the program's own errors */
void *vb_err_stash(void);                      /* the pending error aside (0 if none) ... */
void vb_err_unstash(void *e);                  /* ... and back */
extern VB_IMPORT int (*vb_unhandled_hook)(vb_i32 number, const char *description);
/* Compile On Demand (IDE runs, --debug): a procedure that did not compile reports its compile error when it is
 * called (file:line: error: ... on stderr, as lbvc), then the program ends - unless the hook returns 1 (the robot) */
void vb_compile_error(const char *file, vb_i32 line, const char *message);
extern VB_IMPORT int (*vb_compile_hook)(const char *message);
extern VB_IMPORT const char *(*vb_where_hook)(void);
const char *vb_where(void);                   /* Module.Procedure, line N (a --debug build), or "" */
const char *vb_err_text_pending(void);
#define VB_DBG(l) do { if (vb_dbg_on) { vb_dbg_line(&vb_dbgf, (l)); \
    if (vb_dbgf.jump) { vb_rs = vb_dbgf.jump; vb_dbgf.jump = 0; goto vb_resume_same; } } } while (0)   /* Set Next */
void vb_end(void);                             /* End statement */
/* End: what the runtime does before the program goes (the forms runtime: the controls of the ActiveX control
 * components it shows end, as their own code does in VB6 - the DataRepeater's repeated controls) */
extern VB_IMPORT void (*vb_end_hook)(void);
void vb_stop(void);

#endif
