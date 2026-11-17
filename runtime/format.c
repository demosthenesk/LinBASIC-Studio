/* format.c - Format$, FormatNumber, FormatCurrency, FormatPercent, FormatDateTime,
 * MonthName, WeekdayName (VBA), with the regional settings.
 *
 * Format sections are separated by ';' (positive;negative;zero;null). Numbers: 0 #
 * . , % E+ E- e+ e-; dates: d dd ddd dddd ddddd dddddd w ww m mm mmm mmmm q y yy yyyy
 * h hh n nn s ss ttttt c AM/PM am/pm A/P a/p AMPM : /; text: @ & < > !; \x and "..."
 * are literal. Named formats: General Number, Currency, Fixed, Standard, Percent,
 * Scientific, Yes/No, True/False, On/Off, General Date, Long Date, Medium Date,
 * Short Date, Long Time, Medium Time, Short Time.
 *
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
vbs vbs_sub(vbs s, vb_i32 start, vb_i32 len);

/* ---------------------------------------------------------------- output buffer (UTF-16) */
typedef struct { vb_wchar *p; vb_i32 n, cap; } wb;
static void wput(wb *b, vb_wchar c)
{
    if (b->n == b->cap) {
        vb_i32 nc = b->cap ? b->cap * 2 : 64;
        vb_wchar *np = (vb_wchar *)realloc(b->p, sizeof(vb_wchar) * (size_t)nc);
        if (!np) return;
        b->p = np; b->cap = nc;
    }
    b->p[b->n++] = c;
}
static void wputs(wb *b, const char *utf8)
{
    vb_i32 m = vb_mark();
    vbs s = vb_sfromutf8(utf8);
    for (vb_i32 i = 0; i < s.len; i++) wput(b, s.p[i]);
    vb_release(m);
}
static void wputv(wb *b, vbs s) { for (vb_i32 i = 0; i < s.len; i++) wput(b, s.p[i]); }
static vbs wdone(wb *b) { vbs r = vb_stmp(b->n); if (b->n) memcpy(r.p, b->p, sizeof(vb_wchar) * (size_t)b->n); free(b->p); return r; }

static vb_wchar lower(vb_wchar c) { return c >= 'A' && c <= 'Z' ? c + 32 : c; }
static int ieq(const vb_wchar *p, vb_i32 n, const char *s)
{
    vb_i32 i = 0;
    for (; s[i]; i++) if (i >= n || lower(p[i]) != (vb_wchar)s[i]) return 0;
    return 1;
}
static int is_named(vbs f, const char *name) { return f.len == (vb_i32)strlen(name) && ieq(f.p, f.len, name); }

/* ---------------------------------------------------------------- decimal digits
 * the value with 15 significant digits, rounded half away from zero at `decimals` */
typedef struct { char d[400]; int n; int point; } dec;   /* digits d[0..n), point = digits before '.' */

/* a Decimal being formatted: all its digits (up to 29) instead of the 15 of a Double. g_dval is its value
   as a Double: to_dec uses the digits when v is that value times 10^k (a % multiplies by 100) */
static char g_dsrc[128];
static double g_dval;
static int g_dset;

static void to_dec(double v, int decimals, dec *r)
{
    char buf[64];
    v = fabs(v);
    r->n = 0; r->point = 1;
    if (v == 0) { r->d[0] = '0'; r->n = 1; return; }
    snprintf(buf, sizeof buf, "%.14e", v);
    char dg[64]; int nd = 0; char *p = buf;
    for (; *p && *p != 'e'; p++) if (*p >= '0' && *p <= '9') dg[nd++] = *p;
    int e = atoi(p + 1);
    if (g_dset && g_dval > 0) {
        double ratio = v / g_dval; int k = -1;
        for (int t = 0; t <= 8; t++) { double pw = pow(10, t); if (fabs(ratio - pw) <= pw * 1e-12) { k = t; break; } }
        if (k >= 0) {                              /* the exact digits of the Decimal */
            char all[64]; int na = 0, ip = -1;
            for (const char *q = g_dsrc; *q; q++) { if (*q == '.') ip = na; else if (*q >= '0' && *q <= '9') all[na++] = *q; }
            if (ip < 0) ip = na;
            int f = 0;
            while (f < na - 1 && all[f] == '0') f++;
            int last = na;
            while (last > f + 1 && all[last - 1] == '0') last--;
            nd = 0;
            for (int i = f; i < last; i++) dg[nd++] = all[i];
            e = ip - 1 - f + k;
        }
    }
    /* build with point = e + 1 */
    int point = e + 1;
    int keep = point + decimals;               /* digits to keep */
    if (keep < 0) { r->d[0] = '0'; r->n = 1; r->point = 1; return; }
    char all[400]; int na = 0;
    int lead = point < 1 ? 1 - point : 0;      /* zeros before the first digit */
    for (int i = 0; i < lead && na < 380; i++) all[na++] = '0';
    for (int i = 0; i < nd && na < 380; i++) all[na++] = dg[i];
    int ap = point < 1 ? 1 : point;            /* digits before '.' in all[] */
    while (na < ap + decimals + 1 && na < 390) all[na++] = '0';
    int cut = ap + decimals;
    int up = all[cut] >= '5';
    na = cut;
    if (up) {
        int i = na - 1;
        while (i >= 0 && all[i] == '9') { all[i] = '0'; i--; }
        if (i >= 0) all[i]++;
        else { memmove(all + 1, all, (size_t)na); all[0] = '1'; na++; ap++; }
    }
    /* drop leading zeros of the integer part (keep one) */
    int s = 0;
    while (s < ap - 1 && all[s] == '0') s++;
    memcpy(r->d, all + s, (size_t)(na - s)); r->n = na - s; r->point = ap - s;
}

static int dec_zero(const dec *d) { for (int i = 0; i < d->n; i++) if (d->d[i] != '0') return 0; return 1; }

/* ---------------------------------------------------------------- numbers */
static void put_grouped(wb *b, const char *digits, int n, int group)
{
    for (int i = 0; i < n; i++) {
        wput(b, (vb_wchar)digits[i]);
        int left = n - 1 - i;
        if (group && left > 0 && left % 3 == 0) wput(b, (vb_wchar)(unsigned char)vb_lc_grp);
    }
}

/* one section of a numeric format; sign: 1 = write '-' for a negative value */
static void fmt_num_section(wb *b, double v, const vb_wchar *f, vb_i32 n, int sign)
{
    int int_ph = 0, int_zero = 0, frac_ph = 0, frac_zero = 0, point = 0, group = 0, pct = 0, sci = 0, exp_zero = 0, exp_plus = 0;
    int seen_digit_ph = 0;
    for (vb_i32 i = 0; i < n; i++) {
        vb_wchar c = f[i];
        if (c == '\\') { i++; continue; }
        if (c == '"') { i++; while (i < n && f[i] != '"') i++; continue; }
        if (sci) { if (c == '0' || c == '#') exp_zero += c == '0'; continue; }
        if (c == '0' || c == '#') {
            seen_digit_ph = 1;
            if (point) { frac_ph++; if (c == '0') frac_zero = frac_ph; }
            else { int_ph++; if (c == '0') int_zero++; }
        } else if (c == '.') { if (!point) point = 1; }
        else if (c == ',') { if (!point && seen_digit_ph) group = 1; }
        else if (c == '%') pct++;
        else if ((c == 'E' || c == 'e') && i + 1 < n && (f[i + 1] == '+' || f[i + 1] == '-')) { sci = 1; exp_plus = f[i + 1] == '+'; i++; }
    }
    for (int i = 0; i < pct; i++) v *= 100;
    int neg = v < 0;
    int exponent = 0;
    if (sci && v != 0) {
        int ip = int_ph ? int_ph : 1;
        exponent = (int)floor(log10(fabs(v))) - (ip - 1);
        double m = fabs(v) / pow(10, exponent);
        int ds = g_dset; g_dset = 0;
        dec t; to_dec(m, frac_ph, &t);         /* rounding may carry to one more digit */
        g_dset = ds;
        if (t.point > ip) { exponent++; }
        v = (neg ? -1 : 1) * fabs(v) / pow(10, exponent);
    }
    dec d; to_dec(v, frac_ph, &d);
    /* integer and fraction digit strings */
    char ints[400]; int ni = 0;
    for (int i = 0; i < d.point; i++) ints[ni++] = d.d[i];
    if (ni == 1 && ints[0] == '0') ni = 0;      /* no integer digits: placeholders decide */
    while (ni < int_zero) { memmove(ints + 1, ints, (size_t)ni); ints[0] = '0'; ni++; }
    char fr[400]; int nf = 0;
    for (int i = d.point; i < d.n; i++) fr[nf++] = d.d[i];
    while (nf > frac_zero && fr[nf - 1] == '0') nf--;
    while (nf < frac_zero && nf < (int)sizeof fr) fr[nf++] = '0';   /* Format(0, "0.00"): the zeros all the same */
    if (sign && neg && !(dec_zero(&d))) wput(b, '-');
    int int_done = 0, frac_done = 0;
    for (vb_i32 i = 0; i < n; i++) {
        vb_wchar c = f[i];
        if (c == '\\') { if (i + 1 < n) wput(b, f[++i]); continue; }
        if (c == '"') { i++; while (i < n && f[i] != '"') wput(b, f[i++]); continue; }
        if ((c == 'E' || c == 'e') && sci && i + 1 < n && (f[i + 1] == '+' || f[i + 1] == '-')) {
            if (!int_done) { put_grouped(b, ints, ni, group); int_done = 1; }
            wput(b, c);
            i++;
            if (exponent < 0) wput(b, '-'); else if (exp_plus) wput(b, '+');
            char eb[16]; snprintf(eb, sizeof eb, "%0*d", exp_zero ? exp_zero : 1, exponent < 0 ? -exponent : exponent);
            for (char *q = eb; *q; q++) wput(b, (vb_wchar)*q);
            while (i + 1 < n && (f[i + 1] == '0' || f[i + 1] == '#')) i++;
            continue;
        }
        if (c == '0' || c == '#') {
            if (!int_done && !frac_done) {
                int in_frac = 0;
                for (vb_i32 k = 0; k < i; k++) if (f[k] == '.') in_frac = 1;
                if (!in_frac) { put_grouped(b, ints, ni, group); int_done = 1; continue; }
            }
            if (int_done && !frac_done) {
                int in_frac = 0;
                for (vb_i32 k = 0; k < i; k++) if (f[k] == '.') in_frac = 1;
                if (!in_frac) continue;         /* the other integer placeholders */
            }
            if (!frac_done) { for (int k = 0; k < nf; k++) wput(b, (vb_wchar)fr[k]); frac_done = 1; }
            continue;
        }
        if (c == '.') {
            if (!int_done) { put_grouped(b, ints, ni, group); int_done = 1; }
            wput(b, (vb_wchar)(unsigned char)vb_lc_dec);
            continue;
        }
        if (c == ',') continue;
        wput(b, c);
    }
    if (!int_done && !point && !sci) {
        /* a format without placeholders: just its literal text */
    }
}

/* split f into sections at unquoted ';' */
static int sections(vbs f, vb_i32 *start, vb_i32 *len)
{
    int ns = 0; vb_i32 s = 0;
    for (vb_i32 i = 0; i <= f.len; i++) {
        if (i < f.len && f.p[i] == '\\') { i++; continue; }
        if (i < f.len && f.p[i] == '"') { i++; while (i < f.len && f.p[i] != '"') i++; continue; }
        if (i == f.len || f.p[i] == ';') {
            if (ns < 4) { start[ns] = s; len[ns] = i - s; ns++; }
            s = i + 1;
        }
    }
    return ns;
}

static vbs fmt_number(double v, vbs f)
{
    vb_i32 st[4], ln[4];
    int ns = sections(f, st, ln);
    wb b = { 0, 0, 0 };
    if (v < 0 && ns >= 2 && ln[1] > 0) fmt_num_section(&b, -v, f.p + st[1], ln[1], 0);
    else if (v == 0 && ns >= 3 && ln[2] > 0) fmt_num_section(&b, v, f.p + st[2], ln[2], 1);
    else if (v < 0 && ns >= 2) fmt_num_section(&b, v, f.p + st[0], ln[0], 1);
    else fmt_num_section(&b, v, f.p + st[0], ln[0], 1);
    return wdone(&b);
}

/* number with the regional settings: digits after the point, group, (parens), lead 0 */
static void plain_number(wb *b, double v, int digits, int lead0, int parens, int group)
{
    dec d; to_dec(v, digits, &d);
    int neg = v < 0 && !dec_zero(&d);
    if (neg) wput(b, parens ? '(' : '-');
    char ints[400]; int ni = 0;
    for (int i = 0; i < d.point; i++) ints[ni++] = d.d[i];
    if (ni == 1 && ints[0] == '0' && !lead0 && digits > 0) ni = 0;
    put_grouped(b, ints, ni, group);
    if (digits > 0) {
        wput(b, (vb_wchar)(unsigned char)vb_lc_dec);
        for (int i = d.point; i < d.n; i++) wput(b, (vb_wchar)d.d[i]);
    }
    if (neg && parens) wput(b, ')');
}

static void currency(wb *b, double v, int digits, int lead0, int parens, int group)
{
    const char *pat = v < 0 && !parens ? vb_lc_curneg : vb_lc_curpos;
    if (v < 0 && parens) wput(b, '(');
    for (const char *p = pat; *p; p++) {
        if (*p == 'n') plain_number(b, fabs(v), digits, lead0, 0, group);
        else if (*p == '$') wputs(b, vb_lc_cursym);
        else wput(b, (vb_wchar)(unsigned char)*p);
    }
    if (v < 0 && parens) wput(b, ')');
}

/* ---------------------------------------------------------------- dates */
static int has_ampm(const vb_wchar *f, vb_i32 n)
{
    for (vb_i32 i = 0; i < n; i++) {
        if (f[i] == '"') { i++; while (i < n && f[i] != '"') i++; continue; }
        if (f[i] == '\\') { i++; continue; }
        if (ieq(f + i, n - i, "am/pm") || ieq(f + i, n - i, "a/p") || ieq(f + i, n - i, "ampm")) return 1;
    }
    return 0;
}

static void put_pic(wb *b, const char *pic, double d)
{
    char o[256];
    vb_fmt_pic(o, sizeof o, pic, d);
    wputs(b, o);
}

static void put_num(wb *b, int v, int width)
{
    char t[16]; snprintf(t, sizeof t, width == 2 ? "%02d" : "%d", v);
    for (char *q = t; *q; q++) wput(b, (vb_wchar)*q);
}

static void general_date(wb *b, double d)
{
    vb_i32 m = vb_mark();
    wputv(b, vb_cstr_date(d));
    vb_release(m);
}

static vbs fmt_date(double d, vbs fs, vb_i32 firstday)
{
    const vb_wchar *f = fs.p; vb_i32 n = fs.len;
    wb b = { 0, 0, 0 };
    int y = vb_year(d), mo = vb_month(d), dd = vb_day(d), h = vb_hour(d), mi = vb_minute(d), s = vb_second(d);
    int ampm = has_ampm(f, n);
    int last_h = 0;                             /* m after h or hh is the minute */
    if (firstday < 1) firstday = 1;
    for (vb_i32 i = 0; i < n;) {
        vb_wchar c = lower(f[i]);
        if (f[i] == '\\') { if (i + 1 < n) wput(&b, f[i + 1]); i += 2; continue; }
        if (f[i] == '"') { i++; while (i < n && f[i] != '"') wput(&b, f[i++]); i++; continue; }
        int k = 1;
        while (i + k < n && lower(f[i + k]) == c) k++;
        if (ieq(f + i, n - i, "am/pm")) { wputs(&b, h < 12 ? (f[i] == 'a' ? "am" : "AM") : (f[i] == 'a' ? "pm" : "PM")); i += 5; continue; }
        if (ieq(f + i, n - i, "a/p")) { wput(&b, h < 12 ? (f[i] == 'a' ? 'a' : 'A') : (f[i] == 'a' ? 'p' : 'P')); i += 3; continue; }
        if (ieq(f + i, n - i, "ampm")) { wputs(&b, h < 12 ? vb_lc_am : vb_lc_pm); i += 4; continue; }
        switch (c) {
        case 'd':
            if (k >= 6) put_pic(&b, vb_lc_ldate, d);
            else if (k == 5) put_pic(&b, vb_lc_sdate, d);
            else if (k == 4) wputs(&b, vb_lc_day(vb_weekday(d, 1) - 1, 0));
            else if (k == 3) wputs(&b, vb_lc_day(vb_weekday(d, 1) - 1, 1));
            else put_num(&b, dd, k);
            if (k > 6) k = 6;
            break;
        case 'w':
            if (k >= 2) { put_num(&b, vb_datepart(vb_sascii("ww"), d, firstday, 1), 1); k = 2; }
            else put_num(&b, vb_weekday(d, firstday), 1);
            break;
        case 'm':
            if (last_h && k <= 2) { put_num(&b, mi, k); break; }
            if (k >= 4) { wputs(&b, vb_lc_month(mo, 0)); k = 4; }
            else if (k == 3) wputs(&b, vb_lc_month(mo, 1));
            else put_num(&b, mo, k);
            break;
        case 'n': put_num(&b, mi, k >= 2 ? 2 : 1); if (k > 2) k = 2; break;
        case 's': put_num(&b, s, k >= 2 ? 2 : 1); if (k > 2) k = 2; break;
        case 'q': put_num(&b, (mo - 1) / 3 + 1, 1); k = 1; break;
        case 'y':
            if (k >= 4) { put_num(&b, y, 1); k = 4; }
            else if (k >= 2) { put_num(&b, y % 100, 2); k = 2; }
            else put_num(&b, vb_datepart(vb_sascii("y"), d, 1, 1), 1);
            break;
        case 'h': {
            int hh = ampm ? (h % 12 == 0 ? 12 : h % 12) : h;
            put_num(&b, hh, k >= 2 ? 2 : 1); if (k > 2) k = 2;
            last_h = 2;
            break;
        }
        case 't':
            if (k >= 5) { put_pic(&b, vb_lc_ltime, d); k = 5; }
            else for (int j = 0; j < k; j++) wput(&b, f[i + j]);
            break;
        case 'c': general_date(&b, d); k = 1; break;
        case ':': wput(&b, ':'); break;
        case '/': wput(&b, '/'); break;
        default: for (int j = 0; j < k; j++) wput(&b, f[i + j]);
        }
        if (c != 'h' && c != ':' && c != ' ' && c != '.' && c != '/') last_h = last_h ? last_h - 1 : 0;
        if (c == 'h') last_h = 2;
        i += k;
    }
    return wdone(&b);
}

/* ---------------------------------------------------------------- text */
static vbs fmt_text(vbs s, vbs f)
{
    int nph = 0, left = 0, up = 0, down = 0;
    for (vb_i32 i = 0; i < f.len; i++) {
        vb_wchar c = f.p[i];
        if (c == '@' || c == '&') nph++;
        else if (c == '!') left = 1;
        else if (c == '>') up = 1;
        else if (c == '<') down = 1;
    }
    if (up) s = vb_ucase(s);
    else if (down) s = vb_lcase(s);
    wb b = { 0, 0, 0 };
    if (nph == 0) { wputv(&b, s); return wdone(&b); }
    /* placeholders take the characters from the right (or from the left with !) */
    vb_i32 extra = s.len > nph ? s.len - nph : 0;
    vb_i32 si = left ? 0 : s.len - (s.len < nph ? s.len : nph) - extra;
    int ph = 0;
    vb_i32 missing = nph > s.len ? nph - s.len : 0;
    for (vb_i32 i = 0; i < f.len; i++) {
        vb_wchar c = f.p[i];
        if (c == '!' || c == '<' || c == '>') continue;
        if (c == '\\') { if (i + 1 < f.len) wput(&b, f.p[++i]); continue; }
        if (c == '"') { i++; while (i < f.len && f.p[i] != '"') wput(&b, f.p[i++]); continue; }
        if (c == '@' || c == '&') {
            if (ph == 0 && !left) { for (vb_i32 k = 0; k < extra; k++) wput(&b, s.p[k]); si = extra; }
            int empty = left ? ph >= s.len : ph < missing;
            if (empty) { if (c == '@') wput(&b, ' '); }
            else wput(&b, s.p[si++]);
            ph++;
            if (left && ph == nph) while (si < s.len) wput(&b, s.p[si++]);
            continue;
        }
        wput(&b, c);
    }
    return wdone(&b);
}

/* ---------------------------------------------------------------- Format$ */
static int is_date_format(vbs f)
{
    int date = 0;
    for (vb_i32 i = 0; i < f.len; i++) {
        vb_wchar c = lower(f.p[i]);
        if (f.p[i] == '\\') { i++; continue; }
        if (f.p[i] == '"') { i++; while (i < f.len && f.p[i] != '"') i++; continue; }
        if (c == '0' || c == '#' || c == '@' || c == '&') return 0;
        if (c == 'd' || c == 'm' || c == 'y' || c == 'h' || c == 'n' || c == 's' || c == 'w' || c == 'q' || c == 't' || c == 'c') date = 1;
    }
    return date;
}

static int is_text_format(vbs f)
{
    for (vb_i32 i = 0; i < f.len; i++) {
        vb_wchar c = f.p[i];
        if (c == '\\') { i++; continue; }
        if (c == '"') { i++; while (i < f.len && f.p[i] != '"') i++; continue; }
        if (c == '@' || c == '&' || c == '<' || c == '>' || c == '!') return 1;
        if (c == '0' || c == '#') return 0;
    }
    return 0;
}

/* Yes/No, True/False, On/Off in the language of the regional settings (VB6 on a Greek
 * Windows, confirmed by t_format: Ναι/Όχι, Αληθές/Ψευδές, Ισχύει/Άκυρο) */
static const char *yesno(int yes, int which)
{
    static const char *en[3][2] = { { "No", "Yes" }, { "False", "True" }, { "Off", "On" } };
    static const char *el[3][2] = { { "Όχι", "Ναι" }, { "Ψευδές", "Αληθές" }, { "Άκυρο", "Ισχύει" } };
    return vb_lc_names == 1 ? el[which][yes] : en[which][yes];
}

static vbs named_number(double v, vbs f, int *ok)
{
    wb b = { 0, 0, 0 };
    *ok = 1;
    if (is_named(f, "general number")) { vb_i32 m = vb_mark(); wputv(&b, vb_cstr_d(v)); vb_release(m); }
    else if (is_named(f, "currency")) currency(&b, v, 2, 1, 0, 1);
    else if (is_named(f, "fixed")) plain_number(&b, v, 2, 1, 0, 0);
    else if (is_named(f, "standard")) plain_number(&b, v, 2, 1, 0, 1);
    else if (is_named(f, "percent")) { plain_number(&b, v * 100, 2, 1, 0, 0); wput(&b, '%'); }
    else if (is_named(f, "scientific")) { free(b.p); static const vb_wchar sf[] = { '0', '.', '0', '0', 'E', '+', '0', '0' }; return fmt_number(v, vb_slit(sf, 8)); }
    else if (is_named(f, "yes/no")) wputs(&b, yesno(v != 0, 0));
    else if (is_named(f, "true/false")) wputs(&b, yesno(v != 0, 1));
    else if (is_named(f, "on/off")) wputs(&b, yesno(v != 0, 2));
    else { *ok = 0; free(b.p); return vb_stmp(0); }
    return wdone(&b);
}

static vbs named_date(double d, vbs f, int *ok)
{
    wb b = { 0, 0, 0 };
    *ok = 1;
    if (is_named(f, "general date")) general_date(&b, d);
    else if (is_named(f, "long date")) put_pic(&b, vb_lc_ldate, d);
    else if (is_named(f, "medium date")) put_pic(&b, "dd-MMM-yy", d);
    else if (is_named(f, "short date")) put_pic(&b, vb_lc_sdate, d);
    else if (is_named(f, "long time")) put_pic(&b, vb_lc_ltime, d);
    else if (is_named(f, "medium time")) put_pic(&b, "hh:mm tt", d);
    else if (is_named(f, "short time")) put_pic(&b, "HH:mm", d);
    else { *ok = 0; free(b.p); return vb_stmp(0); }
    return wdone(&b);
}

static vbs format_impl(vbvar v, vbs fmt, vb_i32 has_fmt, vb_i32 firstday, vb_i32 firstweek);
static void dec_source(vbvar v)
{
    char t[100];
    vb_dec_text(v, 0, t);
    snprintf(g_dsrc, sizeof g_dsrc, "%s", t[0] == '-' ? t + 1 : t);
    g_dval = fabs(vb_dec2d(v)); g_dset = 1;
}
vbs vb_format(vbvar v, vbs fmt, vb_i32 has_fmt, vb_i32 firstday, vb_i32 firstweek)
{
    if (v.vt != VT_DECIMAL) return format_impl(v, fmt, has_fmt, firstday, firstweek);
    dec_source(v);
    vbs r = format_impl(v, fmt, has_fmt, firstday, firstweek);
    g_dset = 0;
    return r;
}
static vbs format_impl(vbvar v, vbs fmt, vb_i32 has_fmt, vb_i32 firstday, vb_i32 firstweek)
{
    (void)firstweek;
    if (v.vt == VT_DISPATCH) { v = vb_late(v, "", VBM_GET, 0, 0); if (vb_errpend) return vb_stmp(0); }
    if (v.vt == VT_NULL) {
        if (has_fmt) { vb_i32 st[4], ln[4]; int ns = sections(fmt, st, ln); if (ns >= 4) return fmt_text(vb_stmp(0), vbs_sub(fmt, st[3], ln[3])); }
        return vb_stmp(0);
    }
    if (!has_fmt || fmt.len == 0) {
        if (v.vt == VT_EMPTY) return vb_stmp(0);
        return vb_v2s(v);
    }
    int ok;
    if (v.vt == VT_BSTR) {
        if (is_text_format(fmt)) return fmt_text(v.u.s, fmt);
        int save = vb_errpend;
        double d = vb_s2d(v.u.s);
        if (vb_errpend && !save) {              /* not a number: a date, or the text itself */
            vb_errpend = 0;
            double dt = vb_s2date(v.u.s);
            if (vb_errpend && !save) { vb_errpend = 0; return fmt_text(v.u.s, fmt); }
            v = vb_v_date(dt);
        } else v = vb_v_d(d);
    }
    double x = v.vt == VT_EMPTY ? 0 : v.vt == VT_DATE ? v.u.d : vb_v2d(v);
    if (vb_errpend) return vb_stmp(0);
    vbs r = named_date(x, fmt, &ok);            /* named formats first: "Yes/No" is not a date */
    if (ok) return r;
    r = named_number(x, fmt, &ok);
    if (ok) return r;
    if (v.vt == VT_DATE || is_date_format(fmt)) return fmt_date(x, fmt, firstday);
    if (is_text_format(fmt)) return fmt_text(vb_v2s(v), fmt);
    return fmt_number(x, fmt);
}

vbs vbs_sub(vbs s, vb_i32 start, vb_i32 len) { vbs r = { s.p + start, len, 0 }; return r; }

/* FormatNumber / FormatCurrency / FormatPercent: -1 = True, 0 = False, -2 = the regional setting */
vbs vb_formatnumber(vbvar v, vb_i32 digits, vb_i32 lead0, vb_i32 parens, vb_i32 group, vb_i32 kind)
{
    if (v.vt == VT_DECIMAL) dec_source(v);
    double x = vb_v2d(v);
    if (vb_errpend) return vb_stmp(0);
    if (digits < 0) digits = 2;
    int l0 = lead0 == -2 ? vb_lc_lzero : lead0 != 0;
    int pa = parens == -2 ? (kind == 1 && vb_lc_curneg[0] == '(') : parens != 0;
    int gr = group == -2 ? 1 : group != 0;
    wb b = { 0, 0, 0 };
    if (kind == 1) currency(&b, x, digits, l0, pa, gr);
    else if (kind == 2) { plain_number(&b, x * 100, digits, l0, pa, gr); wput(&b, '%'); }
    else plain_number(&b, x, digits, l0, pa, gr);
    g_dset = 0;
    return wdone(&b);
}

vbs vb_formatdatetime(vb_date d, vb_i32 named)
{
    wb b = { 0, 0, 0 };
    switch (named) {
    case 1: put_pic(&b, vb_lc_ldate, d); break;
    case 2: put_pic(&b, vb_lc_sdate, d); break;
    case 3: put_pic(&b, vb_lc_ltime, d); break;
    case 4: put_pic(&b, "HH:mm", d); break;
    default: general_date(&b, d);
    }
    return wdone(&b);
}

vbs vb_monthname(vb_i32 m, vb_i32 abbrev)
{
    if (m < 1 || m > 12) { vb_raise(5); return vb_stmp(0); }
    return vb_sfromutf8(vb_lc_month(m, abbrev ? 1 : 0));
}

vbs vb_weekdayname(vb_i32 w, vb_i32 abbrev, vb_i32 firstday)
{
    if (w < 1 || w > 7) { vb_raise(5); return vb_stmp(0); }
    if (firstday < 1) firstday = vb_lc_firstday;   /* vbUseSystemDayOfWeek: Monday in Greek (t_format) */
    return vb_sfromutf8(vb_lc_day(w - 1 + firstday - 1, abbrev));
}
