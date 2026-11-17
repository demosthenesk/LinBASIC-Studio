/* lib.c - the VBA library: strings, math, Rnd, dates, interaction.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#define _GNU_SOURCE
#include "vbrt_int.h"
#include <stdint.h>
#include <time.h>
#include <sys/time.h>
#include <math.h>

/* ---------------------------------------------------------------- strings */
vb_i32 vb_len(vbs s) { return s.len; }
static vbs sub(vbs s, vb_i32 from, vb_i32 n)
{
    if (from < 0) from = 0;
    if (from > s.len) from = s.len;
    if (n < 0 || from + n > s.len) n = s.len - from;
    vbs r = vb_stmp(n);
    if (n) memcpy(r.p, s.p + from, sizeof(vb_wchar) * n);
    return r;
}
/* on a byte string (an odd number of bytes) Left, Right and Mid count two bytes a character from the start,
   and the last byte goes with "the rest" (VB6: Mid(3 bytes, 2) = the third byte) */
vbs vb_left(vbs s, vb_i32 n)
{
    if (n < 0) { vb_raise(5); return vb_stmp(0); }
    if (s.flags & VBS_ODD) return vb_leftb(s, n > 0x3FFFFFFF ? 0x7FFFFFFF : 2 * n);
    return sub(s, 0, n);
}
vbs vb_right(vbs s, vb_i32 n)
{
    if (n < 0) { vb_raise(5); return vb_stmp(0); }
    if (s.flags & VBS_ODD) return vb_rightb(s, n > 0x3FFFFFFF ? 0x7FFFFFFF : 2 * n);
    return n >= s.len ? sub(s, 0, s.len) : sub(s, s.len - n, n);
}
vbs vb_mid(vbs s, vb_i32 start, vb_i32 n)
{
    if (start < 1) { vb_raise(5); return vb_stmp(0); }
    if (s.flags & VBS_ODD)
        return start > 0x3FFFFFFF ? vb_stmp(0) : vb_midb(s, 2 * start - 1, n < 0 ? -1 : n > 0x3FFFFFFF ? 0x7FFFFFFF : 2 * n);
    return sub(s, start - 1, n);
}
void vb_mid_stmt(vbs *dst, vb_i32 start, vb_i32 n, vbs src)
{
    if (src.flags & VBS_ODD) src.len++;         /* the last byte with the 0 after it: a character (VB6) */
    if (start < 1 || start > dst->len) { vb_raise(5); return; }
    vb_i32 room = dst->len - (start - 1);
    if (n < 0 || n > src.len) n = src.len;
    if (n > room) n = room;
    vbs copy = sub(*dst, 0, dst->len);
    memcpy(copy.p + start - 1, src.p, sizeof(vb_wchar) * n);
    vb_sset(dst, copy);
}

/* LSet / RSet s = v: v in the length of s, padded with spaces (or cut) */
void vb_slset(vbs *dst, vbs src, vb_i32 right)
{
    vb_i32 n = dst->len;
    vbs r = vb_stmp(n);
    for (vb_i32 i = 0; i < n; i++) r.p[i] = ' ';
    vb_i32 k = src.len < n ? src.len : n;
    if (right) memcpy(r.p + (n - k), src.p, sizeof(vb_wchar) * (size_t)k);
    else memcpy(r.p, src.p, sizeof(vb_wchar) * (size_t)k);
    vb_sset(dst, r);
}

static int eqc(vb_wchar a, vb_wchar b, int text)
{
    if (a == b) return 1;
    if (!text) return 0;
    return vb_wlower(a) == vb_wlower(b);        /* (no temporaries: InStr over a long text) */
}

vb_i32 vb_instr(vb_i32 start, vbs a, vbs b, vb_i32 text)
{
    if (start < 1) { vb_raise(5); return 0; }
    if (start > a.len) return b.len == 0 && start == a.len + 1 ? 0 : 0;
    if (b.len == 0) return start;
    for (vb_i32 i = start - 1; i + b.len <= a.len; i++) {
        vb_i32 k = 0;
        while (k < b.len && eqc(a.p[i + k], b.p[k], text)) k++;
        if (k == b.len) return i + 1;
    }
    return 0;
}
vb_i32 vb_instrrev(vbs a, vbs b, vb_i32 start, vb_i32 text)
{
    if (start == -1 || start > a.len) start = a.len;
    if (start < 1) return 0;
    if (b.len == 0) return start;
    for (vb_i32 i = start - b.len; i >= 0; i--) {
        vb_i32 k = 0;
        while (k < b.len && eqc(a.p[i + k], b.p[k], text)) k++;
        if (k == b.len) return i + 1;
    }
    return 0;
}

vbs vb_ltrim(vbs s) { vb_i32 i = 0; while (i < s.len && s.p[i] == ' ') i++; return sub(s, i, -1); }
vbs vb_rtrim(vbs s) { vb_i32 n = s.len; while (n > 0 && s.p[n - 1] == ' ') n--; return sub(s, 0, n); }
vbs vb_trim(vbs s) { return vb_ltrim(vb_rtrim(s)); }
vbs vb_space(vb_i32 n) { if (n < 0 || n > 0x3FFFFFFF) { vb_raise(5); return vb_stmp(0); } vbs r = vb_stmp(n); for (vb_i32 i = 0; i < r.len; i++) r.p[i] = ' '; return r; }
vbs vb_string_n(vb_i32 n, vb_i32 ch)
{
    if (n < 0 || n > 0x3FFFFFFF) { vb_raise(5); return vb_stmp(0); }   /* (more bytes than a BSTR has) */
    vb_wchar c = ch < 256 && ch >= 0 ? vb_ansi2w((unsigned char)ch) : (vb_wchar)ch;
    vbs r = vb_stmp(n); for (vb_i32 i = 0; i < r.len; i++) r.p[i] = c; return r;
}
vbs vb_string_s(vb_i32 n, vbs ch) { if (ch.len == 0) { vb_raise(5); return vb_stmp(0); } if (n < 0 || n > 0x3FFFFFFF) { vb_raise(5); return vb_stmp(0); } vbs r = vb_stmp(n); for (vb_i32 i = 0; i < r.len; i++) r.p[i] = ch.p[0]; return r; }
vbs vb_chr(vb_i32 code)
{
    if (vb_cp_isdbcs(vb_codepage) && code >= -32768 && code <= 65535 && (code < 0 || code > 255)) {
        unsigned char b[2] = { (unsigned char)((code >> 8) & 0xFF), (unsigned char)(code & 0xFF) };   /* a double-byte code */
        vbs r = vb_stmp(2); r.len = vb_cp_decode(vb_codepage, b, 2, r.p); r.p[r.len] = 0; return r;
    }
    if (code < 0 || code > 255) { vb_raise(5); return vb_stmp(0); }
    vbs r = vb_stmp(1); r.p[0] = vb_ansi2w((unsigned char)code); return r;
}
vbs vb_chrw(vb_i32 code) { if (code < -32768 || code > 65535) { vb_raise(5); return vb_stmp(0); } vbs r = vb_stmp(1); r.p[0] = (vb_wchar)code; return r; }
vb_i32 vb_asc(vbs s)
{
    if (!s.len) { vb_raise(5); return 0; }
    int c = vb_cp_code(vb_codepage, s.p[0], 1);           /* (best fit, as WideCharToMultiByte) */
    if (c < 0) return '?';
    return c > 0xFF ? (vb_i16)c : c;                       /* a double-byte code: an Integer */
}
vb_i32 vb_ascw(vbs s) { if (!s.len) { vb_raise(5); return 0; } return (vb_i16)s.p[0]; }
vbs vb_strreverse(vbs s) { vbs r = vb_stmp(s.len); for (vb_i32 i = 0; i < s.len; i++) r.p[i] = s.p[s.len - 1 - i]; return r; }

vbs vb_replace(vbs s, vbs find, vbs with, vb_i32 start, vb_i32 count, vb_i32 text)
{
    if (start < 1) { vb_raise(5); return vb_stmp(0); }
    vbs src = sub(s, start - 1, -1);                 /* VB returns the string from start on */
    if (find.len == 0 || count == 0) return src;
    vb_i32 cap = src.len + 16, n = 0, done = 0;
    vb_wchar *buf = (vb_wchar *)malloc(sizeof(vb_wchar) * cap);
    for (vb_i32 i = 0; i < src.len;) {
        vb_i32 k = 0;
        if ((count < 0 || done < count) && i + find.len <= src.len)
            while (k < find.len && eqc(src.p[i + k], find.p[k], text)) k++;
        if (k == find.len && (count < 0 || done < count) && i + find.len <= src.len) {
            if (n + with.len >= cap) { cap = (n + with.len) * 2 + 16; buf = (vb_wchar *)realloc(buf, sizeof(vb_wchar) * cap); }
            memcpy(buf + n, with.p, sizeof(vb_wchar) * with.len); n += with.len; i += find.len; done++;
        } else {
            if (n + 1 >= cap) { cap = cap * 2; buf = (vb_wchar *)realloc(buf, sizeof(vb_wchar) * cap); }
            buf[n++] = src.p[i++];
        }
    }
    vbs r = vb_stmp(n); memcpy(r.p, buf, sizeof(vb_wchar) * n); free(buf);
    return r;
}

vbs vb_hex(vb_i64 v, vb_i32 bytes)
{
    char b[24];
    unsigned long long u = (unsigned long long)v;
    if (v < 0) u &= bytes <= 2 ? 0xFFFFULL : 0xFFFFFFFFULL;
    snprintf(b, sizeof b, "%llX", u);
    return vb_sascii(b);
}
vbs vb_oct(vb_i64 v, vb_i32 bytes)
{
    char b[32];
    unsigned long long u = (unsigned long long)v;
    if (v < 0) u &= bytes <= 2 ? 0xFFFFULL : 0xFFFFFFFFULL;
    snprintf(b, sizeof b, "%llo", u);
    return vb_sascii(b);
}
vb_i32 vb_strcomp(vbs a, vbs b, vb_i32 text) { return vb_scmp(a, b, text); }

/* ---------------------------------------------------------------- math */
static uint32_t rnd_seed = 0x50000;
static float rnd_last;
static int rnd_any;
double vb_rnd(double nd, vb_i32 has_n)
{
    float n = has_n ? (float)nd : 1.0f;
    if (n == 0) return rnd_any ? rnd_last : (float)rnd_seed / 16777216.0f;
    if (n < 0) {
        uint32_t m; memcpy(&m, &n, 4);
        rnd_seed = ((m & 0xFFFFFF) + ((m >> 24) & 0xFF)) & 0xFFFFFF;
    }
    rnd_seed = (rnd_seed * 16598013u + 12820163u) & 0xFFFFFF;
    rnd_last = (float)rnd_seed / 16777216.0f;
    rnd_any = 1;
    return rnd_last;
}
void vb_randomize(double seed, vb_i32 has_seed)
{
    if (!has_seed) seed = vb_timer();
    uint64_t b; memcpy(&b, &seed, 8);
    uint32_t hi = (uint32_t)(b >> 32);
    uint32_t h = (hi ^ (hi >> 16)) & 0xFFFF;
    rnd_seed = (rnd_seed & 0xFF) | (h << 8);
}
double vb_timer(void)
{
    struct timeval tv; gettimeofday(&tv, NULL);
    time_t t = tv.tv_sec; struct tm tm; localtime_r(&t, &tm);
    return (float)(tm.tm_hour * 3600 + tm.tm_min * 60 + tm.tm_sec + tv.tv_usec / 1e6);
}
double vb_sgn(double v) { return v > 0 ? 1 : v < 0 ? -1 : 0; }
/* without the floating-point error checks (native code of VB6) they are the instructions of the FPU: NaN, Inf */
double vb_sqr(double v) { if (v < 0 && vb_fpchecks) { vb_raise(5); return 0; } return sqrt(v); }
double vb_log(double v) { if (v <= 0 && vb_fpchecks) { vb_raise(5); return 0; } return log(v); }
double vb_exp(double v) { double r = exp(v); return vb_fpchecks ? vb_chkd(r) : r; }
double vb_sin(double v) { return sin(v); }
double vb_cos(double v) { return cos(v); }
double vb_tan(double v) { return tan(v); }
double vb_atn(double v) { return atan(v); }

/* ---------------------------------------------------------------- dates
 * A Date is the number of days since 30 Dec 1899; the fraction is the time. */
static long days_from_civil(long y, unsigned m, unsigned d)
{
    y -= m <= 2;
    long era = (y >= 0 ? y : y - 399) / 400;
    unsigned yoe = (unsigned)(y - era * 400);
    unsigned doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
    unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + (long)doe - 719468;           /* days since 1970-01-01 */
}
static void civil_from_days(long z, int *y, int *m, int *d)
{
    z += 719468;
    long era = (z >= 0 ? z : z - 146096) / 146097;
    unsigned doe = (unsigned)(z - era * 146097);
    unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    long yy = (long)yoe + era * 400;
    unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    unsigned mp = (5 * doy + 2) / 153;
    *d = (int)(doy - (153 * mp + 2) / 5 + 1);
    *m = (int)(mp < 10 ? mp + 3 : mp - 9);
    *y = (int)(yy + (*m <= 2));
}
#define EPOCH_1899 (-25569L)                       /* 1899-12-30 relative to 1970-01-01 */

vb_date vb_dateserial(vb_i32 y, vb_i32 m, vb_i32 d)
{
    if (y >= 0 && y < 30) y += 2000; else if (y >= 30 && y < 100) y += 1900;
    long ym = (long)y * 12 + (m - 1);
    long yy = ym >= 0 ? ym / 12 : -((-ym + 11) / 12);
    long mm = ym - yy * 12 + 1;
    double r = (double)(days_from_civil(yy, (unsigned)mm, 1) - EPOCH_1899 + (d - 1));
    if (r < -657434.0 || r > 2958465.0) { vb_raise(5); return 0; }   /* outside 1/1/100 .. 12/31/9999 */
    return r;
}
vb_date vb_timeserial(vb_i32 h, vb_i32 m, vb_i32 s) { return ((double)h * 3600 + (double)m * 60 + s) / 86400.0; }

static void split(vb_date v, int *y, int *mo, int *d, int *h, int *mi, int *s)
{
    double day = floor(v);
    double frac = fabs(v - (v < 0 ? ceil(v) : day));
    if (v < 0) day = ceil(v);
    long secs = (long)floor(frac * 86400.0 + 0.5);
    if (secs >= 86400) { secs -= 86400; day += v < 0 ? -1 : 1; }
    civil_from_days((long)day + EPOCH_1899, y, mo, d);
    *h = (int)(secs / 3600); *mi = (int)(secs / 60 % 60); *s = (int)(secs % 60);
}
vb_i32 vb_year(vb_date d) { int y, mo, dd, h, mi, s; split(d, &y, &mo, &dd, &h, &mi, &s); return y; }
vb_i32 vb_month(vb_date d) { int y, mo, dd, h, mi, s; split(d, &y, &mo, &dd, &h, &mi, &s); return mo; }
vb_i32 vb_day(vb_date d) { int y, mo, dd, h, mi, s; split(d, &y, &mo, &dd, &h, &mi, &s); return dd; }
vb_i32 vb_hour(vb_date d) { int y, mo, dd, h, mi, s; split(d, &y, &mo, &dd, &h, &mi, &s); return h; }
vb_i32 vb_minute(vb_date d) { int y, mo, dd, h, mi, s; split(d, &y, &mo, &dd, &h, &mi, &s); return mi; }
vb_i32 vb_second(vb_date d) { int y, mo, dd, h, mi, s; split(d, &y, &mo, &dd, &h, &mi, &s); return s; }
vb_i32 vb_weekday(vb_date d, vb_i32 firstday)
{
    long day = (long)floor(d);
    int w = (int)(((day + 6) % 7 + 7) % 7) + 1;         /* 1 = Sunday (30 Dec 1899 was a Saturday) */
    if (firstday < 1) firstday = 1;
    return ((w - firstday + 7) % 7) + 1;
}
vb_date vb_now(void)
{
    struct timeval tv; gettimeofday(&tv, NULL);
    time_t t = tv.tv_sec; struct tm tm; localtime_r(&t, &tm);
    return vb_dateserial(tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday) + vb_timeserial(tm.tm_hour, tm.tm_min, tm.tm_sec);
}

/* month and day names of the regional settings (Greek, otherwise English) */
static const char *en_month[12] = { "January", "February", "March", "April", "May", "June", "July", "August",
                                     "September", "October", "November", "December" };
static const char *en_mon[12] = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
static const char *en_day[7] = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
static const char *en_dy[7] = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
static const char *el_month[12] = { "Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος", "Ιούλιος",
                                    "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος" };
static const char *el_month_gen[12] = { "Ιανουαρίου", "Φεβρουαρίου", "Μαρτίου", "Απριλίου", "Μαΐου", "Ιουνίου",
                                        "Ιουλίου", "Αυγούστου", "Σεπτεμβρίου", "Οκτωβρίου", "Νοεμβρίου", "Δεκεμβρίου" };
static const char *el_mon[12] = { "Ιαν", "Φεβ", "Μαρ", "Απρ", "Μαϊ", "Ιουν", "Ιουλ", "Αυγ", "Σεπ", "Οκτ", "Νοε", "Δεκ" };
static const char *el_day[7] = { "Κυριακή", "Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο" };
static const char *el_dy[7] = { "Κυρ", "Δευ", "Τρι", "Τετ", "Πεμ", "Παρ", "Σαβ" };

const char *vb_lc_month(int m, int kind)
{
    if (m < 1 || m > 12) return "";
    if (vb_lc_names == 1) return kind == 1 ? el_mon[m - 1] : kind == 2 ? el_month_gen[m - 1] : el_month[m - 1];
    return kind == 1 ? en_mon[m - 1] : en_month[m - 1];
}
const char *vb_lc_day(int wd, int abbrev)
{
    wd = ((wd % 7) + 7) % 7;
    if (vb_lc_names == 1) return abbrev ? el_dy[wd] : el_day[wd];
    return abbrev ? en_dy[wd] : en_day[wd];
}

/* Date pictures of the regional settings: d dd ddd dddd M MM MMM MMMM yy yyyy h hh H HH m mm
 * s ss tt; other characters are copied ('...' quotes literal text). Unlike Windows, VB6
 * keeps the month name in the nominative next to a day number (Greek, t_format). */
static int fmt_pic(char *o, int cap, const char *pic, int y, int mo, int d, int h, int mi, int s, int wd)
{
    int k = 0;
    int genitive = 0;                           /* VB6 Format writes "2 Ιανουάριος": no genitive (t_format) */
    #define PUT(...) do { if (k < cap) k += snprintf(o + k, cap - k, __VA_ARGS__); } while (0)
    for (const char *p = pic; *p;) {
        char c = *p; int n = 1;
        while (p[n] == c) n++;
        switch (c) {
        case 'd': if (n >= 4) PUT("%s", vb_lc_day(wd, 0)); else if (n == 3) PUT("%s", vb_lc_day(wd, 1)); else PUT(n >= 2 ? "%02d" : "%d", d); break;
        case 'M': if (n >= 4) PUT("%s", vb_lc_month(mo, genitive ? 2 : 0)); else if (n == 3) PUT("%s", vb_lc_month(mo, 1)); else PUT(n >= 2 ? "%02d" : "%d", mo); break;
        case 'y': if (n >= 3) PUT("%d", y); else PUT(n == 2 ? "%02d" : "%d", y % 100); break;
        case 'h': PUT(n >= 2 ? "%02d" : "%d", h % 12 == 0 ? 12 : h % 12); break;
        case 'H': PUT(n >= 2 ? "%02d" : "%d", h); break;
        case 'm': PUT(n >= 2 ? "%02d" : "%d", mi); break;
        case 's': PUT(n >= 2 ? "%02d" : "%d", s); break;
        case 't': PUT("%s", h < 12 ? vb_lc_am : vb_lc_pm); break;
        case '\'': { const char *q = p + 1; while (*q && *q != '\'') { PUT("%c", *q); q++; } p = *q ? q + 1 : q; continue; }
        default: for (int i = 0; i < n; i++) PUT("%c", c);
        }
        p += n;
    }
    #undef PUT
    if (k >= cap) k = cap - 1;
    /* no AM/PM symbol in this locale: drop the blank left before or after "tt" */
    while (k > 0 && o[k - 1] == ' ') o[--k] = 0;
    while (o[0] == ' ') memmove(o, o + 1, k--);
    return k;
}

int vb_fmt_pic(char *o, int cap, const char *pic, double date)
{
    int y, mo, d, h, mi, s;
    split(date, &y, &mo, &d, &h, &mi, &s);
    return fmt_pic(o, cap, pic, y, mo, d, h, mi, s, vb_weekday(date, 1) - 1);
}

/* CStr(date), Print, &: short date + long time of the regional settings. VB6 (t_dates):
 * a whole day shows only the date; a value whose day part is 0 (30 Dec 1899) shows only
 * the time, even midnight: CStr(CDate(0)) = "12:00:00 AM". */
vbs vb_cstr_date(vb_date v)
{
    int y, mo, d, h, mi, s; char b[128]; int k = 0;
    split(v, &y, &mo, &d, &h, &mi, &s);
    int has_date = (long)(v < 0 ? ceil(v) : floor(v)) != 0, has_time = (h || mi || s);
    if (has_date) k += fmt_pic(b, sizeof b, vb_lc_sdate, y, mo, d, h, mi, s, 0);
    if (!has_date || has_time) {
        if (k) b[k++] = ' ';
        k += fmt_pic(b + k, (int)sizeof b - k, vb_lc_ltime, y, mo, d, h, mi, s, 0);
    }
    b[k] = 0;
    return vb_sfromutf8(b);
}

/* Write #: #yyyy-mm-dd#, #yyyy-mm-dd hh:mm:ss#; a time alone keeps the 1899-12-30 day. */
void vb_date_iso(vb_date v, char *out)
{
    int y, mo, d, h, mi, s;
    split(v, &y, &mo, &d, &h, &mi, &s);
    int k = sprintf(out, "#%04d-%02d-%02d", y, mo, d);
    if (h || mi || s) k += sprintf(out + k, " %02d:%02d:%02d", h, mi, s);
    strcpy(out + k, "#");
}

/* CDate(text) (oleaut32 VarDateFromStr, simplified): numbers separated by / - . , or blanks
 * are the date, numbers joined by ':' the time; English month names; AM/PM and the
 * symbols of the regional settings. The order of day, month and year comes from the short
 * date of the regional settings; a 4-digit first number means year-month-day; when the
 * order gives an impossible date, day and month are swapped (VB6 accepts 15/3/2010 and
 * 3/15/2010 on any locale). */
static int is_word(const vb_wchar *p, int n, const char *utf8)
{
    if (!*utf8) return 0;
    vb_i32 m = vb_mark();
    vbs w = vb_sfromutf8(utf8);
    int ok = w.len == n;
    for (int i = 0; ok && i < n; i++) {
        vb_wchar a = p[i], b = w.p[i];
        if (a >= 'a' && a <= 'z') a -= 32;
        if (b >= 'a' && b <= 'z') b -= 32;
        if (a != b) ok = 0;
    }
    vb_release(m);
    return ok;
}

static int valid_ymd(long y, long m, long d)
{
    static const int dim[] = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
    if (y < 100) y += y < 30 ? 2000 : 1900;
    if (y < 100 || y > 9999 || m < 1 || m > 12 || d < 1) return 0;
    int leap = (y % 4 == 0 && y % 100 != 0) || y % 400 == 0;
    return d <= dim[m - 1] + (m == 2 && leap);
}

vb_date vb_s2date(vbs s)
{
    static const char *months[] = { "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec" };
    long dn[3]; int dl[3], nd = 0;           /* date numbers and their digit counts */
    long tn[3]; int nt = 0;                  /* time numbers */
    int month_name = 0, ampm = 0, bad = 0, in_time = 0;
    vb_i32 i = 0, n = s.len;
    while (i < n && !bad) {
        vb_wchar c = s.p[i];
        if (c >= '0' && c <= '9') {
            long v = 0; int digits = 0;
            while (i < n && s.p[i] >= '0' && s.p[i] <= '9') { if (digits < 9) v = v * 10 + (s.p[i] - '0'); digits++; i++; }
            if (i < n && s.p[i] == ':') in_time = 1;
            if (in_time) { if (nt < 3) tn[nt++] = v; else bad = 1; }
            else if (nd < 3) { dn[nd] = v; dl[nd++] = digits; } else bad = 1;
            if (in_time && !(i < n && s.p[i] == ':')) in_time = 0;
            if (i < n && s.p[i] == ':') i++;
            continue;
        }
        if (c == ' ' || c == '/' || c == '-' || c == '.' || c == ',' || c == '\t') { i++; continue; }
        vb_i32 j = i;
        while (j < n && s.p[j] != ' ' && s.p[j] != '/' && s.p[j] != '-' && s.p[j] != '.' && s.p[j] != ',' && !(s.p[j] >= '0' && s.p[j] <= '9')) j++;
        const vb_wchar *w = s.p + i; int wl = j - i;
        if (is_word(w, wl, "AM") || is_word(w, wl, "A") || is_word(w, wl, vb_lc_am)) ampm = 1;
        else if (is_word(w, wl, "PM") || is_word(w, wl, "P") || is_word(w, wl, vb_lc_pm)) ampm = 2;
        else {
            int found = 0;
            if (wl >= 3)
                for (int k = 0; k < 12 && !found; k++) {
                    int ok = 1;
                    for (int q = 0; q < 3 && ok; q++) { vb_wchar a = w[q]; if (a >= 'A' && a <= 'Z') a += 32; if (a != months[k][q]) ok = 0; }
                    if (ok && !month_name && nd < 3) { month_name = k + 1; found = 1; }
                }
            if (!found) bad = 1;
        }
        i = j;
    }
    if (bad || (nd == 0 && nt == 0 && !month_name)) { vb_raise(13); return 0; }

    double day = 0;
    if (nd || month_name) {
        long y = -1, m = -1, d = -1;
        /* order of the short date of the regional settings */
        char ord[3]; int no = 0;
        for (const char *p = vb_lc_sdate; *p && no < 3; p++)
            if ((*p == 'd' || *p == 'M' || *p == 'y') && (no == 0 || ord[no - 1] != *p)) ord[no++] = *p;
        if (month_name) {
            m = month_name;
            if (nd == 2) { if (dl[0] >= 3 || dn[0] > 31) { y = dn[0]; d = dn[1]; } else { d = dn[0]; y = dn[1]; } }
            else if (nd == 1) { d = dn[0]; }
            else { vb_raise(13); return 0; }
        } else if (nd == 3) {
            if (dl[0] >= 3 || dn[0] > 31) { y = dn[0]; m = dn[1]; d = dn[2]; }
            else {
                long *slot[3]; int u = 0;
                for (int k = 0; k < no; k++) slot[u++] = ord[k] == 'd' ? &d : ord[k] == 'M' ? &m : &y;
                if (u < 3) { slot[0] = &m; slot[1] = &d; slot[2] = &y; }
                *slot[0] = dn[0]; *slot[1] = dn[1]; *slot[2] = dn[2];
                if (y < 0) y = dn[2];
            }
            if (!valid_ymd(y, m, d) && valid_ymd(y, d, m)) { long t = m; m = d; d = t; }
        } else if (nd == 2) {
            int dfirst = 1;
            for (int k = 0; k < no; k++) if (ord[k] == 'd' || ord[k] == 'M') { dfirst = ord[k] == 'd'; break; }
            if (dl[0] >= 3 || dn[0] > 31) { y = dn[0]; m = dn[1]; d = 1; }
            else if (dfirst) { d = dn[0]; m = dn[1]; } else { m = dn[0]; d = dn[1]; }
            if (y < 0 && !valid_ymd(2000, m, d) && valid_ymd(2000, d, m)) { long t = m; m = d; d = t; }
        } else { vb_raise(13); return 0; }
        if (y < 0) { int cy, cm, cd, h, mi, sc; split(vb_now(), &cy, &cm, &cd, &h, &mi, &sc); y = cy; }
        if (y < 100) y += y < 30 ? 2000 : 1900;
        if (!valid_ymd(y, m, d)) { vb_raise(13); return 0; }
        day = vb_dateserial((vb_i32)y, (vb_i32)m, (vb_i32)d);
    }
    double t = 0;
    if (nt || ampm) {
        long h = nt > 0 ? tn[0] : 0, mi = nt > 1 ? tn[1] : 0, sc = nt > 2 ? tn[2] : 0;
        if (!nt) { vb_raise(13); return 0; }
        if (ampm) { if (h < 1 || h > 12) { vb_raise(13); return 0; } if (h == 12) h = 0; if (ampm == 2) h += 12; }
        if (h > 23 || mi > 59 || sc > 59) { vb_raise(13); return 0; }
        t = vb_timeserial((vb_i32)h, (vb_i32)mi, (vb_i32)sc);
    }
    return day < 0 ? day - t : day + t;
}

/* ---------------------------------------------------------------- interaction (console until the forms runtime) */
vb_i32 (*vb_msgbox_hook)(vbs prompt, vb_i32 buttons, vbs title);
vbs (*vb_inputbox_hook)(vbs prompt, vbs title, vbs def);
vb_i32 vb_msgbox(vbs prompt, vb_i32 buttons, vbs title)
{
    if (vb_msgbox_hook) return vb_msgbox_hook(prompt, buttons, title.len ? title : vb_app_exename());
    vb_i32 m = vb_mark();
    fflush(stdout);
    if (title.len) fprintf(stderr, "[%s] ", vb_stoutf8(title));
    fprintf(stderr, "%s\n", vb_stoutf8(prompt));
    vb_release(m);
    return 1;                                          /* vbOK */
}
vbs vb_inputbox(vbs prompt, vbs title, vbs def)
{
    if (vb_inputbox_hook) return vb_inputbox_hook(prompt, title.len ? title : vb_app_exename(), def);
    vb_i32 m = vb_mark();
    if (title.len) fprintf(stderr, "[%s] ", vb_stoutf8(title));
    fprintf(stderr, "%s ", vb_stoutf8(prompt));
    vb_release(m);
    char buf[4096];
    if (!fgets(buf, sizeof buf, stdin)) return def;
    size_t n = strlen(buf); while (n && (buf[n - 1] == '\n' || buf[n - 1] == '\r')) buf[--n] = 0;
    return vb_sfromutf8(buf);
}
void vb_beep(void) { fputc('\a', stderr); }
void (*vb_sendkeys_hook)(vbs keys, vb_i32 wait);
vbobj (*vb_loadpicture_hook)(vbs path, vb_i32 has_path);
void (*vb_savepicture_hook)(vbobj pic, vbs path);
vbobj vb_loadpicture(vbs path, vb_i32 has_path)
{
    if (vb_loadpicture_hook) return vb_loadpicture_hook(path, has_path);
    vb_raise(481);                                      /* Invalid picture: no forms runtime */
    return 0;
}
void vb_savepicture(vbobj pic, vbs path)
{
    if (vb_savepicture_hook) vb_savepicture_hook(pic, path);
    else vb_raise(481);
}
void (*vb_appactivate_hook)(vbs title);
void vb_appactivate(vbs title, vb_i32 wait)
{
    (void)wait;
    if (vb_appactivate_hook) vb_appactivate_hook(title);   /* a form of the program; other programs: nothing */
}
/* the financial functions: Pmt, IPmt, PPmt, FV, PV, NPer (as VBA computes them) */
static double fin_pow(double r, double n) { return pow(1 + r, n); }
double vb_pmt(double rate, double nper, double pv, double fv, vb_i32 type)
{
    if (nper == 0) { vb_raise(5); return 0; }
    if (rate == 0) return -(pv + fv) / nper;
    double p = fin_pow(rate, nper);
    return -(fv + pv * p) * rate / ((1 + rate * (type ? 1 : 0)) * (p - 1));
}
double vb_fv(double rate, double nper, double pmt, double pv, vb_i32 type)
{
    if (rate == 0) return -(pv + pmt * nper);
    double p = fin_pow(rate, nper);
    return -(pv * p + pmt * (1 + rate * (type ? 1 : 0)) * (p - 1) / rate);
}
double vb_pv(double rate, double nper, double pmt, double fv, vb_i32 type)
{
    if (rate == 0) return -(fv + pmt * nper);
    double p = fin_pow(rate, nper);
    return -(fv + pmt * (1 + rate * (type ? 1 : 0)) * (p - 1) / rate) / p;
}
double vb_ipmt(double rate, double per, double nper, double pv, double fv, vb_i32 type)
{
    if (per < 1 || per > nper) { vb_raise(5); return 0; }
    double pmt = vb_pmt(rate, nper, pv, fv, type);
    if (type && per == 1) return 0;
    double bal = vb_fv(rate, per - 1, pmt, pv, type);
    return type ? (bal - pmt) * rate : bal * rate;
}
double vb_ppmt(double rate, double per, double nper, double pv, double fv, vb_i32 type)
{
    return vb_pmt(rate, nper, pv, fv, type) - vb_ipmt(rate, per, nper, pv, fv, type);
}
double vb_nper(double rate, double pmt, double pv, double fv, vb_i32 type)
{
    if (rate == 0) { if (pmt == 0) { vb_raise(5); return 0; } return -(pv + fv) / pmt; }
    double k = pmt * (1 + rate * (type ? 1 : 0)) / rate;
    double a = (k - fv), b = (k + pv);
    if (a / b <= 0) { vb_raise(5); return 0; }
    return log(a / b) / log(1 + rate);
}
/* DDB, SLN, SYD, NPV, IRR, MIRR, Rate: the algorithms of VBA (as Microsoft.VisualBasic.Financial has them) */
double vb_sln(double cost, double salvage, double life)
{
    if (life == 0) { vb_raise(5); return 0; }
    return (cost - salvage) / life;
}
double vb_syd(double cost, double salvage, double life, double period)
{
    if (salvage < 0 || period <= 0 || period > life) { vb_raise(5); return 0; }
    double r = (cost - salvage) / (life * (life + 1));
    return r * (life + 1 - period) * 2;
}
double vb_ddb(double cost, double salvage, double life, double period, double factor)
{
    if (factor <= 0 || salvage < 0 || period <= 0 || period > life) { vb_raise(5); return 0; }
    if (cost <= 0) return 0;
    if (life < 2) return cost - salvage;
    if (life == 2 && period > 1) return 0;
    if (life == 2 && period <= 1) return cost - salvage;
    if (period <= 1) {
        double r = cost * factor / life, t = cost - salvage;
        return r > t ? t : r;
    }
    double t = (life - factor) / life, n1 = period - 1;
    double r = factor * cost / life * pow(t, n1);
    double tot = cost * (1 - pow(t, n1));
    double excess = tot + r - cost + salvage;
    if (excess > 0) r -= excess;
    return r >= 0 ? r : 0;
}
/* the elements of an array argument as Doubles (a Double array, or numbers in another one) */
static double *fin_values(const vbarr *a, vb_i32 vt, vb_i32 *n)
{
    *n = 0;
    if (!a || !a->nd || a->n <= 0) { vb_raise(5); return 0; }
    double *d = (double *)malloc(sizeof(double) * (size_t)a->n);
    for (vb_i32 i = 0; i < a->n; i++) {
        const char *e = a->data + (size_t)i * (size_t)a->esz;
        switch (vt) {
        case VT_R8: memcpy(&d[i], e, 8); break;
        case VT_R4: { float f; memcpy(&f, e, 4); d[i] = f; break; }
        case VT_I4: { vb_i32 v; memcpy(&v, e, 4); d[i] = v; break; }
        case VT_I2: case VT_BOOL: { vb_i16 v; memcpy(&v, e, 2); d[i] = v; break; }
        case VT_UI1: d[i] = (unsigned char)*e; break;
        case VT_CY: { vb_cy v; memcpy(&v, e, 8); d[i] = (double)v / 10000.0; break; }
        case VT_VARIANT: d[i] = vb_v2d(*(const vbvar *)e); break;
        default: free(d); vb_raise(13); return 0;
        }
        if (vb_errpend) { free(d); return 0; }
    }
    *n = a->n;
    return d;
}
/* sign: 0 every value, -1 only the payments (< 0), 1 only the receipts (> 0) */
static double fin_npv(double rate, const double *v, vb_i32 n, int sign)
{
    double tot = 0, disc = 1;
    for (vb_i32 i = 0; i < n; i++) {
        disc += disc * rate;
        if (!((sign < 0 && v[i] > 0) || (sign > 0 && v[i] < 0))) tot += v[i] / disc;
    }
    return tot;
}
double vb_npv(double rate, const vbarr *a, vb_i32 vt)
{
    vb_i32 n;
    if (rate == -1) { vb_raise(5); return 0; }
    double *v = fin_values(a, vt, &n);
    if (!v) return 0;
    double r = fin_npv(rate, v, n, 0);
    free(v);
    return r;
}
static double fin_pv2(const double *v, vb_i32 n, double guess)
{
    double tot = 0, t = 1 + guess;
    vb_i32 lo = 0;
    while (lo < n && v[lo] == 0) lo++;
    for (vb_i32 i = n - 1; i >= lo; i--) { tot /= t; tot += v[i]; }
    return tot;
}
double vb_irr(const vbarr *a, vb_i32 vt, double guess)
{
    vb_i32 n;
    if (guess <= -1) { vb_raise(5); return 0; }
    double *v = fin_values(a, vt, &n);
    if (!v) return 0;
    if (n <= 1) { free(v); vb_raise(5); return 0; }
    double t = v[0] > 0 ? v[0] : -v[0];
    for (vb_i32 i = 0; i < n; i++) { if (v[i] > t) t = v[i]; else if (-v[i] > t) t = -v[i]; }
    double eps = t * 1e-7 * 0.01, r0 = guess, y0 = fin_pv2(v, n, r0), r1, y1;
    r1 = y0 > 0 ? r0 + 1e-5 : r0 - 1e-5;
    if (r1 <= -1) { free(v); vb_raise(5); return 0; }
    y1 = fin_pv2(v, n, r1);
    for (int k = 0; k <= 39; k++) {
        if (y1 == y0) {
            r0 = r1 > r0 ? r0 - 1e-5 : r0 + 1e-5;
            y0 = fin_pv2(v, n, r0);
            if (y1 == y0) break;
        }
        r0 = r1 - (r1 - r0) * y1 / (y1 - y0);
        if (r0 <= -1) r0 = (r1 - 1) * 0.5;
        y0 = fin_pv2(v, n, r0);
        if (fabs(y0) < eps) { free(v); return r0; }
        double tmp = y0; y0 = y1; y1 = tmp;
        tmp = r0; r0 = r1; r1 = tmp;
    }
    free(v);
    vb_raise(5);                                /* no result after the iterations */
    return 0;
}
double vb_mirr(const vbarr *a, vb_i32 vt, double finance, double reinvest)
{
    vb_i32 n;
    if (finance == -1 || reinvest == -1) { vb_raise(5); return 0; }
    double *v = fin_values(a, vt, &n);
    if (!v) return 0;
    if (n <= 1) { free(v); vb_raise(5); return 0; }
    double neg = fin_npv(finance, v, n, -1);
    if (neg == 0) { free(v); vb_raise(11); return 0; }      /* no payment: Division by zero */
    double pos = fin_npv(reinvest, v, n, 1);
    free(v);
    double t = -pos * pow(reinvest + 1, (double)n) / (neg * (finance + 1));
    if (t < 0) { vb_raise(5); return 0; }
    return pow(t, 1.0 / (n - 1)) - 1;
}
static double fin_rate_eval(double rate, double nper, double pmt, double pv, double fv, vb_i32 due)
{
    if (rate == 0) return pv + pmt * nper + fv;
    double t = pow(rate + 1, nper), t2 = due ? 1 + rate : 1;
    return pv * t + pmt * t2 * (t - 1) / rate + fv;
}
double vb_rate(double nper, double pmt, double pv, double fv, vb_i32 due, double guess)
{
    if (nper <= 0) { vb_raise(5); return 0; }
    double r0 = guess, y0 = fin_rate_eval(r0, nper, pmt, pv, fv, due);
    double r1 = y0 > 0 ? r0 / 2 : r0 * 2, y1 = fin_rate_eval(r1, nper, pmt, pv, fv, due);
    for (int k = 0; k <= 39; k++) {
        if (y1 == y0) {
            r0 = r1 > r0 ? r0 - 1e-5 : r0 + 1e-5;
            y0 = fin_rate_eval(r0, nper, pmt, pv, fv, due);
            if (y1 == y0) break;
        }
        r0 = r1 - (r1 - r0) * y1 / (y1 - y0);
        y0 = fin_rate_eval(r0, nper, pmt, pv, fv, due);
        if (fabs(y0) < 1e-7) return r0;
        double tmp = y0; y0 = y1; y1 = tmp;
        tmp = r0; r0 = r1; r1 = tmp;
    }
    vb_raise(5);
    return 0;
}
/* Partition(Number, Start, Stop, Interval): "lower:upper", each as wide as the wider of Stop + 1 and Start - 1 */
vbs vb_partition(vb_i32 number, vb_i32 start, vb_i32 stop, vb_i32 interval)
{
    if (start < 0 || stop <= start || interval < 1) { vb_raise(5); return vb_stmp(0); }
    vb_i32 lower = 0, upper = 0;
    int nolower = 0, noupper = 0;
    if (number < start) { upper = start - 1; nolower = 1; }
    else if (number > stop) { lower = stop + 1; noupper = 1; }
    else if (interval == 1) { lower = upper = number; }
    else {
        lower = ((number - start) / interval) * interval + start;
        upper = lower + interval - 1;
        if (upper > stop) upper = stop;
        if (lower < start) lower = start;
    }
    char b1[32], b2[32], out[80];
    int w1 = snprintf(b1, sizeof b1, "%ld", (long)stop + 1), w2 = snprintf(b2, sizeof b2, "%ld", (long)start - 1);
    int sp = w1 > w2 ? w1 : w2;
    if (nolower) { int w = snprintf(b1, sizeof b1, "%ld", (long)upper); if (w > sp) sp = w; }
    int k = 0;
    if (nolower) k += snprintf(out + k, sizeof out - (size_t)k, "%*s", sp, "");
    else k += snprintf(out + k, sizeof out - (size_t)k, "%*ld", sp, (long)lower);
    out[k++] = ':';
    if (noupper) k += snprintf(out + k, sizeof out - (size_t)k, "%*s", sp, "");
    else k += snprintf(out + k, sizeof out - (size_t)k, "%*ld", sp, (long)upper);
    out[k] = 0;
    return vb_sfromutf8(out);
}
void vb_sendkeys(vbs keys, vb_bool wait)
{
    if (vb_sendkeys_hook) vb_sendkeys_hook(keys, wait);
    else vb_raise(5);                                   /* no window to send keys to */
}
vb_i32 vb_rgb(vb_i32 r, vb_i32 g, vb_i32 b)
{
    if (r < 0 || g < 0 || b < 0) { vb_raise(5); return 0; }
    if (r > 255) r = 255;
    if (g > 255) g = 255;
    if (b > 255) b = 255;
    return r | (g << 8) | (b << 16);
}
vb_i32 vb_qbcolor(vb_i32 c)
{
    static const vb_i32 t[16] = { 0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0,
                                  0x808080, 0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF };
    if (c < 0 || c > 15) { vb_raise(5); return 0; }
    return t[c];
}
void (*vb_doevents_hook)(void);
void vb_doevents(void) { fflush(stdout); if (vb_doevents_hook) vb_doevents_hook(); }

/* ---------------------------------------------------------------- Like
 * ? one character, * any characters, # a digit, [list] [!list] with ranges a-z;
 * text = Option Compare Text (no case). */
static vb_wchar like_fold(vb_wchar c, vb_i32 text)
{
    if (!text) return c;
    return vb_wlower(c);
}

static int like_at(const vb_wchar *s, vb_i32 ns, const vb_wchar *p, vb_i32 np, vb_i32 text)
{
    vb_i32 i = 0, j = 0;
    while (j < np) {
        vb_wchar c = p[j];
        if (c == '*') {
            while (j < np && p[j] == '*') j++;
            if (j == np) return 1;
            for (vb_i32 k = i; k <= ns; k++) if (like_at(s + k, ns - k, p + j, np - j, text)) return 1;
            return 0;
        }
        if (i >= ns) return 0;
        if (c == '?') { i++; j++; continue; }
        if (c == '#') { if (s[i] < '0' || s[i] > '9') return 0; i++; j++; continue; }
        if (c == '[') {
            vb_i32 e = j + 1;
            while (e < np && p[e] != ']') e++;
            if (e >= np) { vb_raise(93); return 0; }      /* Invalid pattern string */
            vb_i32 a = j + 1;
            int neg = 0;
            if (a < e && p[a] == '!') { neg = 1; a++; }
            if (a == e && !neg) return 0;                  /* [] matches nothing here */
            int hit = 0;
            vb_wchar ch = like_fold(s[i], text);
            for (vb_i32 k = a; k < e; k++) {
                if (k + 2 < e && p[k + 1] == '-') {
                    vb_wchar lo = like_fold(p[k], text), hi = like_fold(p[k + 2], text);
                    if (ch >= lo && ch <= hi) hit = 1;
                    k += 2;
                } else if (ch == like_fold(p[k], text)) hit = 1;
            }
            if (hit == neg) return 0;
            i++; j = e + 1;
            continue;
        }
        if (like_fold(s[i], text) != like_fold(c, text)) return 0;
        i++; j++;
    }
    return i == ns;
}

vb_bool vb_like(vbs s, vbs pattern, vb_i32 text)
{
    if (pattern.len == 2 && pattern.p[0] == '[' && pattern.p[1] == ']') return s.len == 0 ? -1 : 0;
    return like_at(s.p, s.len, pattern.p, pattern.len, text) ? -1 : 0;
}

/* ---------------------------------------------------------------- DateAdd, DateDiff, DatePart */
static int interval(vbs iv)
{
    vb_i32 m = vb_mark();
    char *u = vb_stoutf8(vb_lcase(vb_trim(iv)));
    static const char *names[] = { "yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s", 0 };
    int r = -1;
    for (int i = 0; names[i]; i++) if (!strcmp(u, names[i])) r = i;
    vb_release(m);
    if (r < 0) vb_raise(5);
    return r;
}
enum { IV_YYYY, IV_Q, IV_M, IV_Y, IV_D, IV_W, IV_WW, IV_H, IV_N, IV_S };

static double date_of(int y, int m, int d, double frac)
{
    double day = vb_dateserial(y, m, 1);
    int dim = (int)(vb_dateserial(y, m + 1, 1) - day);
    if (d > dim) d = dim;                       /* Jan 31 + 1 month = Feb 29 */
    day += d - 1;
    return day < 0 ? day - frac : day + frac;
}

vb_date vb_dateadd(vbs iv, double n, vb_date d)
{
    int k = interval(iv);
    if (k < 0) return 0;
    int y, mo, dd, h, mi, s;
    split(d, &y, &mo, &dd, &h, &mi, &s);
    double frac = fabs(d - (d < 0 ? ceil(d) : floor(d)));
    long num = (long)vb_fix(n);
    switch (k) {
    case IV_YYYY: return date_of(y + (int)num, mo, dd, frac);
    case IV_Q: num *= 3; /* fall through */
    case IV_M: { long t = (long)y * 12 + (mo - 1) + num; int ny = (int)(t >= 0 ? t / 12 : -((-t + 11) / 12)); return date_of(ny, (int)(t - (long)ny * 12) + 1, dd, frac); }
    case IV_Y: case IV_D: case IV_W: return d + num;
    case IV_WW: return d + 7.0 * num;
    case IV_H: return d + num / 24.0;
    case IV_N: return d + num / 1440.0;
    case IV_S: return d + num / 86400.0;
    }
    return d;
}

static long week_start(double day, int firstday)
{
    long dn = (long)floor(day);
    int w = vb_weekday(dn, firstday);           /* 1 .. 7 from firstday */
    return dn - (w - 1);
}

vb_i32 vb_datediff(vbs iv, vb_date a, vb_date b, vb_i32 firstday)
{
    int k = interval(iv);
    if (k < 0) return 0;
    if (firstday < 1) firstday = 1;
    int y1, m1, d1, h1, n1, s1, y2, m2, d2, h2, n2, s2;
    split(a, &y1, &m1, &d1, &h1, &n1, &s1);
    split(b, &y2, &m2, &d2, &h2, &n2, &s2);
    double secs1 = floor(a) * 86400.0 + h1 * 3600 + n1 * 60 + s1, secs2 = floor(b) * 86400.0 + h2 * 3600 + n2 * 60 + s2;
    switch (k) {
    case IV_YYYY: return y2 - y1;
    case IV_Q: return (y2 * 4 + (m2 - 1) / 3) - (y1 * 4 + (m1 - 1) / 3);
    case IV_M: return (y2 * 12 + m2) - (y1 * 12 + m1);
    case IV_Y: case IV_D: return (vb_i32)(floor(b) - floor(a));
    case IV_W: return (vb_i32)((floor(b) - floor(a)) / 7);
    case IV_WW: return (vb_i32)((week_start(b, firstday) - week_start(a, firstday)) / 7);
    case IV_H: return (vb_i32)(floor(secs2 / 3600) - floor(secs1 / 3600));
    case IV_N: return (vb_i32)(floor(secs2 / 60) - floor(secs1 / 60));
    case IV_S: return (vb_i32)(secs2 - secs1);
    }
    return 0;
}

vb_i32 vb_datepart(vbs iv, vb_date d, vb_i32 firstday, vb_i32 firstweek)
{
    int k = interval(iv);
    if (k < 0) return 0;
    if (firstday < 1) firstday = 1;
    int y, mo, dd, h, mi, s;
    split(d, &y, &mo, &dd, &h, &mi, &s);
    (void)firstweek;
    switch (k) {
    case IV_YYYY: return y;
    case IV_Q: return (mo - 1) / 3 + 1;
    case IV_M: return mo;
    case IV_Y: return (vb_i32)(floor(d) - vb_dateserial(y, 1, 1)) + 1;
    case IV_D: return dd;
    case IV_W: return vb_weekday(d, firstday);
    case IV_WW: return (vb_i32)((week_start(d, firstday) - week_start(vb_dateserial(y, 1, 1), firstday)) / 7) + 1;
    case IV_H: return h;
    case IV_N: return mi;
    case IV_S: return s;
    }
    return 0;
}

vb_date vb_datevalue(vbs s) { double d = vb_s2date(s); return d < 0 ? ceil(d) : floor(d); }
vb_date vb_timevalue(vbs s) { double d = vb_s2date(s); return fabs(d - (d < 0 ? ceil(d) : floor(d))); }

/* StrConv: 1 upper, 2 lower, 3 proper case (first letter of each word) */
/* a byte string as VB6 keeps it: two bytes in each character (InputB, StrConv vbFromUnicode) */
/* ---------------------------------------------------------------- the byte functions: bytes, not characters */
vb_i32 vb_lenb(vbs s) { return vb_sbytes(s); }
vbs vb_chrb(vb_i32 code)
{
    if (code < 0 || code > 255) { vb_raise(6); return vb_stmp(0); }
    char c = (char)code;
    return vb_sfrombytes(&c, 1);
}
vb_i32 vb_ascb(vbs s) { if (!vb_sbytes(s)) { vb_raise(5); return 0; } return vb_sbyte(s, 0); }
vbs vb_leftb(vbs s, vb_i32 n)
{
    if (n < 0) { vb_raise(5); return vb_stmp(0); }
    vb_i32 nb = vb_sbytes(s);
    return vb_sfrombytes2(s, 0, n < nb ? n : nb);
}
vbs vb_rightb(vbs s, vb_i32 n)
{
    if (n < 0) { vb_raise(5); return vb_stmp(0); }
    vb_i32 nb = vb_sbytes(s);
    if (n > nb) n = nb;
    return vb_sfrombytes2(s, nb - n, n);
}
vbs vb_midb(vbs s, vb_i32 start, vb_i32 n)      /* n < 0: the rest */
{
    if (start < 1) { vb_raise(5); return vb_stmp(0); }
    vb_i32 nb = vb_sbytes(s);
    if (start > nb) return vb_stmp(0);
    if (n < 0 || start - 1 + n > nb) n = nb - (start - 1);
    return vb_sfrombytes2(s, start - 1, n);
}
static int byte_eq(int a, int b, int text)
{
    if (text) { if (a >= 'A' && a <= 'Z') a += 32; if (b >= 'A' && b <= 'Z') b += 32; }
    return a == b;
}
vb_i32 vb_instrb(vb_i32 start, vbs s, vbs find, vb_i32 text)
{
    (void)text;                                 /* InStrB compares the bytes as they are (VB6, vbTextCompare too) */
    text = 0;
    if (start < 1) { vb_raise(5); return 0; }
    vb_i32 ns = vb_sbytes(s), nf = vb_sbytes(find);
    if (nf == 0) return start <= ns ? start : (ns == 0 && start == 1 ? 1 : 0);
    for (vb_i32 i = start - 1; i + nf <= ns; i++) {
        vb_i32 k = 0;
        while (k < nf && byte_eq(vb_sbyte(s, i + k), vb_sbyte(find, k), text)) k++;
        if (k == nf) return i + 1;
    }
    return 0;
}
/* MidB(s, start[, n]) = value: the bytes are replaced, the length stays */
void vb_midb_stmt(vbs *s, vb_i32 start, vb_i32 n, vbs value)
{
    vb_i32 nb = vb_sbytes(*s);
    if (start < 1 || start > nb) { vb_raise(5); return; }
    vb_i32 nv = vb_sbytes(value);
    if (n < 0 || n > nv) n = nv;
    if (start - 1 + n > nb) n = nb - (start - 1);
    char *b = (char *)malloc((size_t)(nb + 1));
    for (vb_i32 i = 0; i < nb; i++) b[i] = (char)vb_sbyte(*s, i);
    for (vb_i32 i = 0; i < n; i++) b[start - 1 + i] = (char)vb_sbyte(value, i);
    vb_i32 m = vb_mark();
    vb_sset(s, vb_sfrombytes(b, nb));
    vb_release(m);
    free(b);
}

vbs vb_strconv(vbs s, vb_i32 conv) { return vb_strconv_l(s, conv, 0, 0); }
