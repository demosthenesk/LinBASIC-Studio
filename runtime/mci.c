#define _GNU_SOURCE
#pragma GCC diagnostic ignored "-Wformat-truncation"
/* mci.c - the Media Control Interface of Windows (winmm) for Linux.
 *
 * VB6 programs play sound and video through MCI: mciSendString("play x from 0 wait", ...) with
 * Declare, or the Multimedia control (MCI32.OCX), which sends the same commands. On Windows the
 * runtime calls the real winmm; on Linux this file is the MCI: the command strings, the devices
 * (waveaudio, sequencer, avivideo, mpegvideo, cdaudio), the time formats, the notifications and the
 * error numbers and messages of Windows.
 *
 * Lengths and positions come from the files themselves (the WAVE header, the tempo map of a MIDI
 * file, the AVI header), as the MCI drivers of Windows compute them; while a device plays, its
 * position follows the clock. The sound and the pictures are made by GStreamer (LGPL), loaded when
 * a device first plays (dlopen): a computer without it still runs the program, silently.
 *   VBRT_MCI_SINK   the GStreamer audio sink (default: autoaudiosink; "none": no sound at all)
 *
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include <ctype.h>
#include <stdint.h>
#include <limits.h>
#include <math.h>
#include <time.h>
#include "vbos.h"
#ifndef _WIN32
#include <dlfcn.h>
#include <unistd.h>
#include <dirent.h>
#include <fcntl.h>
#include <sys/ioctl.h>
#include <sys/stat.h>
#ifdef __linux__
#include <linux/cdrom.h>
#endif
#endif

void (*vb_mci_post_hook)(vb_i64 hwnd, vb_i32 code, vb_i32 devid);

#ifdef _WIN32
/* ---------------------------------------------------------------- Windows: the real winmm (mci_win.c) */
int lbmci_send(const char *cmd, char *ret, int retlen, void (*cb)(void *, int, int), void *ctx, long long hwnd);
int lbmci_errtext(int code, char *buf, int len);
int lbmci_devid(const char *name);

vb_i32 vb_mci_send_hwnd(const char *cmd, char *ret, vb_i32 retlen, vb_mci_notify cb, void *ctx, vb_i64 hwnd)
{
    return lbmci_send(cmd, ret, retlen, (void (*)(void *, int, int))cb, ctx, hwnd);
}
vb_i32 vb_mci_send(const char *cmd, char *ret, vb_i32 retlen, vb_mci_notify cb, void *ctx)
{
    return lbmci_send(cmd, ret, retlen, (void (*)(void *, int, int))cb, ctx, 0);
}
vb_i32 vb_mci_error_text(vb_i32 code, char *buf, vb_i32 len) { return lbmci_errtext(code, buf, len); }
void lbmci_forget(void *ctx);
void vb_mci_forget(void *ctx) { lbmci_forget(ctx); }
void vb_mci_poll(void) {}
vb_i32 vb_mci_active(void) { return 0; }
vb_i32 vb_mci_next_video(vb_i32 after) { (void)after; return 0; }
int vb_mci_frame(vb_i32 id, vb_mci_frameinfo *f) { (void)id; (void)f; return 0; }
#else

/* ================================================================ Linux */

/* ---------------------------------------------------------------- GStreamer, loaded at run time */
typedef struct { void *memory; int flags; unsigned char *data; size_t size, maxsize; void *user_data[4]; void *res[4]; } gmap;
static struct {
    int tried, ok;
    int (*init_check)(int *, char ***, void **);
    void *(*factory_make)(const char *, const char *);
    int (*set_state)(void *, int);
    int (*get_state)(void *, int *, int *, uint64_t);
    int (*query_position)(void *, int, int64_t *);
    int (*query_duration)(void *, int, int64_t *);
    int (*seek)(void *, double, int, int, int, int64_t, int, int64_t);
    void (*object_unref)(void *);
    void *(*caps_from_string)(const char *);
    void (*mini_unref)(void *);
    void *(*sample_get_buffer)(void *);
    void *(*sample_get_caps)(void *);
    void *(*caps_get_structure)(void *, unsigned);
    int (*structure_get_int)(void *, const char *, int *);
    int (*buffer_map)(void *, gmap *, int);
    void (*buffer_unmap)(void *, gmap *);
    char *(*filename_to_uri)(const char *, void **);
    void *(*element_get_bus)(void *);
    void *(*bus_pop)(void *);
    void (*object_set)(void *, const char *, ...);
    void (*g_free)(void *);
    void *(*try_pull_sample)(void *, uint64_t);
    void *(*try_pull_preroll)(void *, uint64_t);
    void *(*parse_launch)(const char *, void **);
    void *(*bin_get_by_name)(void *, const char *);
} G;

static int gst_load(void)
{
    if (G.tried) return G.ok;
    G.tried = 1;
    const char *sink = getenv("VBRT_MCI_SINK");
    if (sink && !strcmp(sink, "none")) return 0;
    void *g = dlopen("libgstreamer-1.0.so.0", RTLD_NOW | RTLD_GLOBAL);
    void *o = dlopen("libgobject-2.0.so.0", RTLD_NOW | RTLD_GLOBAL);
    void *l = dlopen("libglib-2.0.so.0", RTLD_NOW | RTLD_GLOBAL);
    void *a = dlopen("libgstapp-1.0.so.0", RTLD_NOW | RTLD_GLOBAL);
    if (!g || !o || !l) return 0;
#define S(f, h, n) if (!(*(void **)&G.f = dlsym(h, n))) return 0;
    S(init_check, g, "gst_init_check") S(factory_make, g, "gst_element_factory_make")
    S(set_state, g, "gst_element_set_state") S(get_state, g, "gst_element_get_state")
    S(query_position, g, "gst_element_query_position") S(query_duration, g, "gst_element_query_duration")
    S(seek, g, "gst_element_seek") S(object_unref, g, "gst_object_unref")
    S(caps_from_string, g, "gst_caps_from_string") S(mini_unref, g, "gst_mini_object_unref")
    S(sample_get_buffer, g, "gst_sample_get_buffer") S(sample_get_caps, g, "gst_sample_get_caps")
    S(caps_get_structure, g, "gst_caps_get_structure") S(structure_get_int, g, "gst_structure_get_int")
    S(buffer_map, g, "gst_buffer_map") S(buffer_unmap, g, "gst_buffer_unmap")
    S(filename_to_uri, g, "gst_filename_to_uri") S(element_get_bus, g, "gst_element_get_bus")
    S(bus_pop, g, "gst_bus_pop") S(object_set, o, "g_object_set") S(g_free, l, "g_free")
    S(parse_launch, g, "gst_parse_launch") S(bin_get_by_name, g, "gst_bin_get_by_name")
#undef S
    if (a) {
        *(void **)&G.try_pull_sample = dlsym(a, "gst_app_sink_try_pull_sample");
        *(void **)&G.try_pull_preroll = dlsym(a, "gst_app_sink_try_pull_preroll");
    }
    if (!G.init_check(0, 0, 0)) return 0;
    return G.ok = 1;
}

enum { GST_NULL = 1, GST_READY = 2, GST_PAUSED = 3, GST_PLAYING = 4, GFMT_TIME = 3,
       GS_FLUSH = 1, GS_ACCURATE = 2, GS_NONE = 0, GS_SET = 1 };

typedef struct {
    void *pipe, *vsink, *bus;
    int prerolled;
} gplay;

static gplay *gp_open(const char *path, int video, const char *uri)
{
    if (!gst_load()) return 0;
    void *p = G.factory_make("playbin", 0);
    if (!p) return 0;
    char *u = 0;
    if (uri) G.object_set(p, "uri", uri, (void *)0);
    else {
        u = G.filename_to_uri(path, 0);
        if (!u) { G.object_unref(p); return 0; }
        G.object_set(p, "uri", u, (void *)0);
        G.g_free(u);
    }
    const char *sname = getenv("VBRT_MCI_SINK");
    if (sname && *sname) {
        void *as = G.factory_make(sname, 0);
        if (as) { if (!strcmp(sname, "fakesink")) G.object_set(as, "sync", 1, (void *)0); G.object_set(p, "audio-sink", as, (void *)0); }
    }
    gplay *g = (gplay *)calloc(1, sizeof *g);
    g->pipe = p;
    if (video && G.try_pull_sample) {
        void *vs = G.factory_make("appsink", 0);
        if (vs) {
            void *caps = G.caps_from_string("video/x-raw,format=BGRx");
            G.object_set(vs, "caps", caps, "max-buffers", 1u, "drop", 1, "sync", 1, (void *)0);
            G.mini_unref(caps);
            G.object_set(p, "video-sink", vs, (void *)0);
            g->vsink = vs;
        }
    }
    if (!g->vsink) G.object_set(p, "flags", 2 | 0x10 /* audio, soft volume */, (void *)0);
    g->bus = G.element_get_bus(p);
    G.set_state(p, GST_PAUSED);
    return g;
}
static void gp_wait_state(gplay *g)
{
    int st = 0, pend = 0;
    G.get_state(g->pipe, &st, &pend, 2000000000ull);
}
static void gp_play(gplay *g, double from_ms, double to_ms)
{
    if (!g) return;
    if (to_ms >= 0 && to_ms <= from_ms) return;      /* nothing to play */
    gp_wait_state(g);
    G.seek(g->pipe, 1.0, GFMT_TIME, GS_FLUSH | GS_ACCURATE, GS_SET, (int64_t)(from_ms * 1e6),
           to_ms >= 0 ? GS_SET : GS_NONE, to_ms >= 0 ? (int64_t)(to_ms * 1e6) : -1);
    G.set_state(g->pipe, GST_PLAYING);
}
static void gp_pause(gplay *g) { if (g) G.set_state(g->pipe, GST_PAUSED); }
static void gp_seek(gplay *g, double ms)
{
    if (!g) return;
    G.set_state(g->pipe, GST_PAUSED);
    gp_wait_state(g);
    G.seek(g->pipe, 1.0, GFMT_TIME, GS_FLUSH | GS_ACCURATE, GS_SET, (int64_t)(ms * 1e6), GS_NONE, -1);
}
static void gp_volume(gplay *g, double v) { if (g) G.object_set(g->pipe, "volume", v, (void *)0); }
static double gp_duration(gplay *g)
{
    if (!g) return -1;
    gp_wait_state(g);
    int64_t d = -1;
    for (int i = 0; i < 20; i++) {
        if (G.query_duration(g->pipe, GFMT_TIME, &d) && d >= 0) return d / 1e6;
        usleep(10000);
    }
    return -1;
}
static void gp_drain(gplay *g)
{
    void *m;
    if (g && g->bus) while ((m = G.bus_pop(g->bus))) G.mini_unref(m);
}
static void gp_close(gplay *g)
{
    if (!g) return;
    G.set_state(g->pipe, GST_NULL);
    gp_drain(g);
    if (g->bus) G.object_unref(g->bus);
    G.object_unref(g->pipe);
    free(g);
}

/* ---------------------------------------------------------------- the errors of MCI */
enum {
    E_INVALID_DEVICE_ID = 257, E_UNRECOGNIZED_KEYWORD = 259, E_UNRECOGNIZED_COMMAND = 261, E_HARDWARE = 262,
    E_INVALID_DEVICE_NAME = 263, E_OUT_OF_MEMORY = 264, E_DEVICE_OPEN = 265, E_CANNOT_LOAD_DRIVER = 266,
    E_MISSING_COMMAND_STRING = 267, E_PARAM_OVERFLOW = 268, E_MISSING_STRING_ARGUMENT = 269, E_BAD_INTEGER = 270,
    E_PARSER_INTERNAL = 271, E_DRIVER_INTERNAL = 272, E_MISSING_PARAMETER = 273, E_UNSUPPORTED_FUNCTION = 274,
    E_FILE_NOT_FOUND = 275, E_DEVICE_NOT_READY = 276, E_INTERNAL = 277, E_DRIVER = 278, E_CANNOT_USE_ALL = 279,
    E_MULTIPLE = 280, E_EXTENSION_NOT_FOUND = 281, E_OUTOFRANGE = 282, E_FLAGS_NOT_COMPATIBLE = 284,
    E_FILE_NOT_SAVED = 286, E_DEVICE_TYPE_REQUIRED = 287, E_DEVICE_LOCKED = 288, E_DUPLICATE_ALIAS = 289,
    E_BAD_CONSTANT = 290, E_MUST_USE_SHAREABLE = 291, E_MISSING_DEVICE_NAME = 292, E_BAD_TIME_FORMAT = 293,
    E_NO_CLOSING_QUOTE = 294, E_DUPLICATE_FLAGS = 295, E_INVALID_FILE = 296, E_NULL_PARAMETER_BLOCK = 297,
    E_UNNAMED_RESOURCE = 298, E_NEW_REQUIRES_ALIAS = 299, E_NOTIFY_ON_AUTO_OPEN = 300, E_NO_ELEMENT_ALLOWED = 301,
    E_NONAPPLICABLE_FUNCTION = 302, E_ILLEGAL_FOR_AUTO_OPEN = 303, E_FILENAME_REQUIRED = 304,
    E_EXTRA_CHARACTERS = 305, E_DEVICE_NOT_INSTALLED = 306, E_GET_CD = 307, E_SET_CD = 308, E_SET_DRIVE = 309,
    E_DEVICE_LENGTH = 310, E_DEVICE_ORD_LENGTH = 311, E_NO_INTEGER = 312,
    E_WAVE_OUTPUTSINUSE = 320, E_WAVE_INPUTSUNSUITABLE = 328, E_SEQ_NOMIDIPRESENT = 343,
    E_NO_WINDOW = 346, E_CREATEWINDOW = 347, E_FILE_READ = 348, E_FILE_WRITE = 349, E_NO_IDENTITY = 350
};

#include "mci_errors.inc"           /* mci_errtab: the messages of mciGetErrorString (VB6 oracle d_mci) */

vb_i32 vb_mci_error_text(vb_i32 code, char *buf, vb_i32 len)
{
    for (int i = 0; mci_errtab[i].text; i++)
        if (mci_errtab[i].code == code) {
            if (buf && len > 0) { snprintf(buf, (size_t)len, "%s", mci_errtab[i].text); }
            return 1;
        }
    if (buf && len > 0) buf[0] = 0;
    return 0;
}

/* ---------------------------------------------------------------- devices */
enum { T_WAVE = 1, T_SEQ, T_AVI, T_MPEG, T_CD };
static const char *const tnames[] = { "", "waveaudio", "sequencer", "avivideo", "mpegvideo", "cdaudio" };

/* time formats: the MCI_FORMAT numbers */
enum { TF_MS = 0, TF_HMS = 1, TF_MSF = 2, TF_FRAMES = 3, TF_SMPTE24 = 4, TF_SMPTE25 = 5, TF_SMPTE30 = 6,
       TF_SMPTE30DROP = 7, TF_BYTES = 8, TF_SAMPLES = 9, TF_TMSF = 10, TF_SONGPTR = 0x4001 };

enum { M_NOTREADY = 524, M_STOP = 525, M_PLAY = 526, M_RECORD = 527, M_SEEK = 528, M_PAUSE = 529, M_OPEN = 530 };

typedef struct { int64_t tick; double ms; double us_per_q; } tempo_pt;
typedef struct { int start, len; int audio; } cdtrack;          /* in frames (1/75 s) */

typedef struct mdev {
    int id, type, tf, mode, autoopen, isnew, shareable;
    char alias[256], elem[1024], path[1024];
    /* the native unit of the device: bytes (wave), ticks (sequencer), frames (avi, cd), ms (mpeg) */
    double len, pos, from, to;
    double t0;                                  /* the clock when it started playing */
    int playing, recording;
    /* wave */
    int ch, rate, bits, align, avgbps, ftag;
    unsigned char *rec; size_t nrec, caprec;   /* the samples of the file, with what was recorded */
    size_t recpos;                              /* where the recording goes */
    void *rpipe, *rsink;                        /* the capture (GStreamer) */
    char tmpwav[300];                           /* the recording, for playing it before it is saved */
    /* sequencer */
    int division, ntracks, smpte, midifmt, tempo_moved;
    tempo_pt *tmap; int ntmap;
    /* video */
    double usframe; int vw, vh; vb_i64 hwnd; int winshow; int has_video, has_audio;
    /* cd */
    cdtrack trk[100]; int ntrk; int cdtrack;
    /* notification */
    vb_mci_notify cb; void *ctx; vb_i64 nhwnd; int npending;
    double speed; int audio_off;
    gplay *gp; int gp_failed;
    struct mdev *next;
} mdev;

static mdev *devs;

typedef struct { vb_mci_notify cb; void *ctx; vb_i64 hwnd; int id, code; } nq;
static nq *queue; static int nqueue, capqueue;

static double now_ms(void)
{
    struct timespec ts; clock_gettime(CLOCK_MONOTONIC, &ts);
    return ts.tv_sec * 1000.0 + ts.tv_nsec / 1e6;
}

static void post(mdev *d, int code)
{
    if (!d->npending) return;
    d->npending = 0;
    if (nqueue == capqueue) { capqueue = capqueue ? capqueue * 2 : 16; queue = (nq *)realloc(queue, sizeof(nq) * capqueue); }
    nq q = { d->cb, d->ctx, d->nhwnd, d->id, code };
    queue[nqueue++] = q;
}

static mdev *find_dev(const char *name)
{
    for (mdev *d = devs; d; d = d->next)
        if (!strcasecmp(d->alias, name)) return d;
    return 0;
}

/* ---------------------------------------------------------------- the files */
static unsigned rd16(const unsigned char *p) { return p[0] | p[1] << 8; }
static unsigned rd32(const unsigned char *p) { return p[0] | p[1] << 8 | p[2] << 16 | (unsigned)p[3] << 24; }
static unsigned be16(const unsigned char *p) { return p[0] << 8 | p[1]; }
static unsigned be32(const unsigned char *p) { return (unsigned)p[0] << 24 | p[1] << 16 | p[2] << 8 | p[3]; }

static unsigned char *slurp(const char *path, size_t *n, size_t max)
{
    FILE *f = fopen(path, "rb");
    if (!f) return 0;
    unsigned char *b = (unsigned char *)malloc(max);
    *n = fread(b, 1, max, f);
    fclose(f);
    return b;
}

static int parse_wave(mdev *d)
{
    size_t n;
    unsigned char *b = slurp(d->path, &n, 1 << 16);
    if (!b) return E_FILE_NOT_FOUND;
    int ok = 0;
    if (n >= 12 && !memcmp(b, "RIFF", 4) && !memcmp(b + 8, "WAVE", 4)) {
        FILE *f = fopen(d->path, "rb");
        size_t off = 12;
        unsigned char h[8], fmt[16];
        long long fsize = 0;
        if (f) { fseek(f, 0, SEEK_END); fsize = ftell(f); }
        int gotfmt = 0;
        while (f && off + 8 <= (size_t)fsize) {
            fseek(f, (long)off, SEEK_SET);
            if (fread(h, 1, 8, f) != 8) break;
            unsigned sz = rd32(h + 4);
            if (!memcmp(h, "fmt ", 4) && sz >= 16 && fread(fmt, 1, 16, f) == 16) {
                d->ftag = (int)rd16(fmt); d->ch = (int)rd16(fmt + 2); d->rate = (int)rd32(fmt + 4);
                d->avgbps = (int)rd32(fmt + 8); d->align = (int)rd16(fmt + 12); d->bits = (int)rd16(fmt + 14);
                gotfmt = 1;
            } else if (!memcmp(h, "data", 4) && gotfmt) {
                long long avail = fsize - (long long)off - 8;
                if ((long long)sz > avail) sz = (unsigned)avail;
                d->len = d->align > 0 ? (double)(sz / (unsigned)d->align * (unsigned)d->align) : sz;
                ok = 1;
                break;
            }
            off += 8 + sz + (sz & 1);
        }
        if (f) fclose(f);
    }
    free(b);
    return ok && d->avgbps > 0 ? 0 : E_INVALID_FILE;
}

/* the tempo map of a MIDI file: tick -> ms */
static int parse_midi(mdev *d)
{
    size_t n;
    unsigned char *b = slurp(d->path, &n, 1 << 22);
    if (!b) return E_FILE_NOT_FOUND;
    size_t base = 0;
    if (n >= 20 && !memcmp(b, "RIFF", 4) && !memcmp(b + 8, "RMID", 4)) base = 20;   /* .rmi */
    if (n < base + 14 || memcmp(b + base, "MThd", 4)) { free(b); return E_INVALID_FILE; }
    int ntrk = (int)be16(b + base + 10), div = (int)be16(b + base + 12);
    d->midifmt = (int)be16(b + base + 8);
    d->ntracks = ntrk;
    d->smpte = (div & 0x8000) != 0;
    d->division = div & 0x7FFF;
    if (d->smpte) d->division = (256 - (div >> 8)) * (div & 0xFF);   /* ticks per second */
    if (d->division <= 0) d->division = 96;
    size_t p = base + 8 + be32(b + base + 4);
    typedef struct { int64_t tick; double usq; } tev;
    tev *tv = 0; int ntv = 0, cap = 0;
    int64_t maxtick = 0;
    for (int t = 0; t < ntrk && p + 8 <= n; t++) {
        if (memcmp(b + p, "MTrk", 4)) break;
        size_t len = be32(b + p + 4), q = p + 8, end = p + 8 + len;
        if (end > n) end = n;
        int64_t tick = 0; int run = 0;
        while (q < end) {
            unsigned v = 0; int k = 0;
            do { v = (v << 7) | (b[q] & 0x7F); k++; } while ((b[q++] & 0x80) && q < end && k < 4);
            tick += v;
            if (q >= end) break;
            int st = b[q];
            if (st == 0xFF) {
                int type = b[q + 1]; q += 2;
                unsigned l = 0; do { l = (l << 7) | (b[q] & 0x7F); } while ((b[q++] & 0x80) && q < end);
                if (type == 0x51 && l == 3 && q + 3 <= end) {
                    if (ntv == cap) { cap = cap ? cap * 2 : 16; tv = (tev *)realloc(tv, sizeof(tev) * cap); }
                    tv[ntv].tick = tick; tv[ntv].usq = (b[q] << 16) | (b[q + 1] << 8) | b[q + 2]; ntv++;
                }
                q += l;
                if (type == 0x2F) break;
            } else if (st == 0xF0 || st == 0xF7) {
                q++;
                unsigned l = 0; do { l = (l << 7) | (b[q] & 0x7F); } while ((b[q++] & 0x80) && q < end);
                q += l;
            } else {
                if (st & 0x80) { run = st; q++; }
                int hi = run & 0xF0;
                q += (hi == 0xC0 || hi == 0xD0) ? 1 : 2;
            }
        }
        if (tick > maxtick) maxtick = tick;
        p = end;
    }
    free(b);
    /* sort the tempo events by tick (stable) */
    for (int i = 1; i < ntv; i++) for (int j = i; j > 0 && tv[j - 1].tick > tv[j].tick; j--) { tev x = tv[j]; tv[j] = tv[j - 1]; tv[j - 1] = x; }
    d->tmap = (tempo_pt *)malloc(sizeof(tempo_pt) * (ntv + 1));
    d->ntmap = 0;
    double usq = 500000, ms = 0; int64_t at = 0;
    d->tmap[d->ntmap++] = (tempo_pt){ 0, 0, usq };
    for (int i = 0; i < ntv; i++) {
        if (d->smpte) break;
        ms += (tv[i].tick - at) * usq / d->division / 1000.0;
        at = tv[i].tick; usq = tv[i].usq;
        if (tv[i].tick == d->tmap[d->ntmap - 1].tick) d->tmap[d->ntmap - 1].us_per_q = usq;
        else d->tmap[d->ntmap++] = (tempo_pt){ at, ms, usq };
    }
    free(tv);
    d->len = (double)maxtick;
    return 0;
}
/* ticks <-> ms as MCISEQ does it (VB6 oracles d_mci, d_mci2): whole microseconds per tick in each
 * part of the tempo map; a format 1 file has the last tempo of the file all the way (MCISEQ);
 * ms -> ticks rounds, from the start of the part in whole ms */
static int seq_part(mdev *d, double tick)
{
    if (d->midifmt == 1) return d->ntmap - 1;
    int i = d->ntmap - 1;
    while (i > 0 && d->tmap[i].tick > tick) i--;
    return i;
}
static double part_uspt(mdev *d, int i) { return floor(d->tmap[i].us_per_q / d->division); }
static double part_us(mdev *d, int i)        /* where the part starts, in microseconds */
{
    if (d->midifmt == 1) return 0;
    double us = 0;
    for (int j = 1; j <= i; j++) us += (d->tmap[j].tick - d->tmap[j - 1].tick) * part_uspt(d, j - 1);
    return us;
}
static double seq_ms(mdev *d, double tick)
{
    if (d->smpte) return tick * 1000.0 / d->division;
    int i = seq_part(d, tick);
    double t0 = d->midifmt == 1 ? 0 : (double)d->tmap[i].tick;
    return (part_us(d, i) + (tick - t0) * part_uspt(d, i)) / 1000.0;
}
static double seq_tick(mdev *d, double ms)
{
    if (d->smpte) return ms * d->division / 1000.0;
    int i = d->midifmt == 1 ? d->ntmap - 1 : 0;
    if (d->midifmt != 1)
        while (i + 1 < d->ntmap && ms * 1000.0 >= part_us(d, i + 1)) i++;
    double u = part_uspt(d, i);
    double t0 = d->midifmt == 1 ? 0 : (double)d->tmap[i].tick;
    return u > 0 ? t0 + floor((ms - floor(part_us(d, i) / 1000.0)) * 1000.0 / u + 0.5) : 0;
}
static double seq_tempo(mdev *d, double tick)   /* the tempo at a tick, bpm: the last one until a seek */
{
    int i = d->ntmap - 1;
    if (d->tempo_moved) while (i > 0 && d->tmap[i].tick > tick) i--;
    return 60000000.0 / d->tmap[i].us_per_q;
}

static int parse_avi(mdev *d)
{
    size_t n;
    unsigned char *b = slurp(d->path, &n, 1 << 16);
    if (!b) return E_FILE_NOT_FOUND;
    int ok = 0;
    if (n >= 12 && !memcmp(b, "RIFF", 4) && !memcmp(b + 8, "AVI ", 4)) {
        for (size_t i = 12; i + 64 < n; i++)
            if (!memcmp(b + i, "avih", 4)) {
                const unsigned char *h = b + i + 8;
                d->usframe = rd32(h); d->len = rd32(h + 16); d->vw = (int)rd32(h + 32); d->vh = (int)rd32(h + 36);
                ok = 1;
                break;
            }
        d->has_audio = memmem(b, n, "auds", 4) != 0;
    }
    free(b);
    if (!ok || d->usframe <= 0) return E_INVALID_FILE;
    d->has_video = 1;
    return 0;
}

/* ---------------------------------------------------------------- the CD (Linux ioctls) */
static const char *cd_dev(void)
{
    const char *e = getenv("VBRT_CDROM");
    return e && *e ? e : "/dev/cdrom";
}
static int cd_read_toc(mdev *d)
{
#ifdef __linux__
    int fd = open(cd_dev(), O_RDONLY | O_NONBLOCK);
    if (fd < 0) return E_DEVICE_NOT_READY;
    struct cdrom_tochdr th;
    d->ntrk = 0;
    if (ioctl(fd, CDROM_DRIVE_STATUS, CDSL_CURRENT) != CDS_DISC_OK || ioctl(fd, CDROMREADTOCHDR, &th) != 0) { close(fd); return 0; }
    int first = th.cdth_trk0, last = th.cdth_trk1;
    int prev = -1;
    for (int t = first; t <= last + 1 && d->ntrk < 99; t++) {
        struct cdrom_tocentry te; memset(&te, 0, sizeof te);
        te.cdte_track = t > last ? CDROM_LEADOUT : (unsigned char)t;
        te.cdte_format = CDROM_MSF;
        if (ioctl(fd, CDROMREADTOCENTRY, &te) != 0) break;
        int fr = (te.cdte_addr.msf.minute * 60 + te.cdte_addr.msf.second) * 75 + te.cdte_addr.msf.frame;
        if (prev >= 0) d->trk[d->ntrk - 1].len = fr - prev;
        if (t <= last) { d->trk[d->ntrk].start = fr; d->trk[d->ntrk].audio = !(te.cdte_ctrl & CDROM_DATA_TRACK); d->ntrk++; }
        prev = fr;
    }
    close(fd);
    if (d->ntrk) d->len = d->trk[d->ntrk - 1].start + d->trk[d->ntrk - 1].len - d->trk[0].start;
#endif
    return 0;
}
static int cd_door(int open_it)
{
#ifdef __linux__
    int fd = open(cd_dev(), O_RDONLY | O_NONBLOCK);
    if (fd < 0) return E_DEVICE_NOT_READY;
    int r = ioctl(fd, open_it ? CDROMEJECT : CDROMCLOSETRAY, 0);
    close(fd);
    return r == 0 ? 0 : E_HARDWARE;
#else
    (void)open_it; return E_UNSUPPORTED_FUNCTION;
#endif
}

/* ---------------------------------------------------------------- units and time formats */
static double nat_ms(mdev *d, double n)
{
    switch (d->type) {
    case T_WAVE: return d->avgbps ? n * 1000.0 / d->avgbps : 0;
    case T_SEQ: return seq_ms(d, n);
    case T_AVI: return n * d->usframe / 1000.0;
    case T_CD: return n * 1000.0 / 75;
    }
    return n;
}
static double ms_nat(mdev *d, double ms)
{
    switch (d->type) {
    case T_WAVE: return floor(ms * d->avgbps / 1000.0 / (d->align ? d->align : 1)) * (d->align ? d->align : 1);
    case T_SEQ: return seq_tick(d, ms);
    case T_AVI: return floor(ms * 1000.0 / d->usframe);
    case T_CD: return floor(ms * 75 / 1000.0);
    }
    return ms;
}

/* the position now (native units) */
static void rec_pull(mdev *d);
static double cur_pos(mdev *d)
{
    if (d->recording) {
        rec_pull(d);
        if (d->rpipe) return (double)d->recpos;
        double n = floor((now_ms() - d->t0) * d->avgbps / 1000.0 / (d->align ? d->align : 1)) * (d->align ? d->align : 1);
        return (double)d->recpos + n;
    }
    if (!d->playing) return d->pos;
    double p = ms_nat(d, nat_ms(d, d->from) + (now_ms() - d->t0) * (d->speed > 0 ? d->speed / 1000.0 : 1));
    return p > d->to ? d->to : p;
}

static void out_num(char *r, int n, double v) { snprintf(r, (size_t)n, "%.0f", floor(v + 0.5)); }

/* a position or length in the time format of the device */
static void fmt_time(mdev *d, double nat, char *r, int n)
{
    double ms = nat_ms(d, nat);
    switch (d->tf) {
    case TF_BYTES: out_num(r, n, nat); return;
    case TF_SAMPLES: out_num(r, n, d->align ? nat / d->align : nat); return;
    case TF_FRAMES:
        if (d->type == T_AVI) out_num(r, n, nat);
        else out_num(r, n, floor(ms * (d->usframe > 0 ? 1000.0 / d->usframe : 30) / 1000.0));
        return;
    case TF_SONGPTR: out_num(r, n, floor(nat * 4 / d->division)); return;   /* lengths; positions: fmt_pos */
    case TF_MSF: {
        long fr = d->type == T_CD ? (long)nat + (d->ntrk ? d->trk[0].start : 0) : (long)floor(ms * 75 / 1000.0);
        snprintf(r, (size_t)n, "%02ld:%02ld:%02ld", fr / 4500, fr / 75 % 60, fr % 75);
        return;
    }
    case TF_TMSF: {
        long fr = (long)nat + (d->ntrk ? d->trk[0].start : 0);
        int t = 0;
        while (t + 1 < d->ntrk && d->trk[t + 1].start <= fr) t++;
        long rel = d->ntrk ? fr - d->trk[t].start : fr;
        snprintf(r, (size_t)n, "%02d:%02ld:%02ld:%02ld", t + 1, rel / 4500, rel / 75 % 60, rel % 75);
        return;
    }
    case TF_HMS: { long s = (long)floor(ms / 1000); snprintf(r, (size_t)n, "%02ld:%02ld:%02ld", s / 3600, s / 60 % 60, s % 60); return; }
    case TF_SMPTE24: case TF_SMPTE25: case TF_SMPTE30: case TF_SMPTE30DROP: {
        int fps = d->tf == TF_SMPTE24 ? 24 : d->tf == TF_SMPTE25 ? 25 : 30;
        long fr = (long)floor(ms * fps / 1000.0);
        snprintf(r, (size_t)n, "%02ld:%02ld:%02ld:%02ld", fr / fps / 3600, fr / fps / 60 % 60, fr / fps % 60, fr % fps);
        return;
    }
    }
    out_num(r, n, ms);
}
/* a length: the waveaudio driver rounds up (MulDiv) */
static void fmt_len(mdev *d, double nat, char *r, int n) { fmt_time(d, nat, r, n); }

/* a number of the command in the time format of the device -> native units; -1 if bad */
static int parse_time(mdev *d, const char *s, double *out)
{
    char *e;
    if (d->tf == TF_MSF || d->tf == TF_TMSF || d->tf == TF_HMS || (d->tf >= TF_SMPTE24 && d->tf <= TF_SMPTE30DROP)) {
        long f[4] = { 0, 0, 0, 0 }; int k = 0;
        const char *p = s;
        while (k < 4) {
            f[k++] = strtol(p, &e, 10);
            if (e == p) return E_BAD_INTEGER;
            if (*e != ':') break;
            p = e + 1;
        }
        if (*e) return E_BAD_INTEGER;
        if (d->tf == TF_TMSF) {
            int t = (int)f[0];
            if (t < 1 || t > (d->ntrk ? d->ntrk : 1)) return E_OUTOFRANGE;
            long fr = (f[1] * 60 + f[2]) * 75 + f[3];
            *out = (d->ntrk ? d->trk[t - 1].start - d->trk[0].start : 0) + fr;
            return 0;
        }
        if (d->tf == TF_MSF) { long fr = (f[0] * 60 + f[1]) * 75 + f[2]; *out = d->type == T_CD ? fr - (d->ntrk ? d->trk[0].start : 0) : ms_nat(d, fr * 1000.0 / 75); return 0; }
        if (d->tf == TF_HMS) { *out = ms_nat(d, (f[0] * 3600 + f[1] * 60 + f[2]) * 1000.0); return 0; }
        int fps = d->tf == TF_SMPTE24 ? 24 : d->tf == TF_SMPTE25 ? 25 : 30;
        *out = ms_nat(d, ((f[0] * 3600 + f[1] * 60 + f[2]) * fps + f[3]) * 1000.0 / fps);
        return 0;
    }
    double v = strtod(s, &e);
    if (e == s || *e) return E_BAD_INTEGER;
    if (v < 0) return E_OUTOFRANGE;
    switch (d->tf) {
    case TF_BYTES: *out = d->align ? floor(v / d->align) * d->align : v; return 0;
    case TF_SAMPLES: *out = v * (d->align ? d->align : 1); return 0;
    case TF_FRAMES: *out = d->type == T_AVI ? v : ms_nat(d, v * (d->usframe > 0 ? d->usframe / 1000.0 : 1000.0 / 30)); return 0;
    case TF_SONGPTR: *out = v * d->division / 4; return 0;
    }
    *out = ms_nat(d, v);
    return 0;
}

static const char *tf_name(int tf)
{
    switch (tf) {
    case TF_MS: return "milliseconds"; case TF_HMS: return "hms"; case TF_MSF: return "msf";
    case TF_FRAMES: return "frames"; case TF_SMPTE24: return "smpte 24"; case TF_SMPTE25: return "smpte 25";
    case TF_SMPTE30: return "smpte 30"; case TF_SMPTE30DROP: return "smpte 30 drop"; case TF_BYTES: return "bytes";
    case TF_SAMPLES: return "samples"; case TF_TMSF: return "tmsf"; case TF_SONGPTR: return "song pointer";
    }
    return "";
}

/* ---------------------------------------------------------------- the command line */
#define MAXTOK 32
typedef struct { char *t[MAXTOK]; int n; int used[MAXTOK]; } toks;

static int tokenize(char *s, toks *k)
{
    k->n = 0;
    memset(k->used, 0, sizeof k->used);
    while (*s) {
        while (*s == ' ' || *s == '\t') s++;
        if (!*s) break;
        if (k->n == MAXTOK) return E_PARAM_OVERFLOW;
        if (*s == '"') {
            s++;
            k->t[k->n++] = s;
            while (*s && *s != '"') s++;
            if (!*s) return E_NO_CLOSING_QUOTE;
            *s++ = 0;
        } else {
            k->t[k->n++] = s;
            while (*s && *s != ' ' && *s != '\t') s++;
            if (*s) *s++ = 0;
        }
    }
    return 0;
}
/* the flag word(s) at token i: "time", "format" ... */
static int kw(toks *k, int i, const char *w)
{
    return i < k->n && !strcasecmp(k->t[i], w);
}
static int find_kw(toks *k, int from, const char *w)
{
    for (int i = from; i < k->n; i++) if (!k->used[i] && !strcasecmp(k->t[i], w)) return i;
    return -1;
}

/* ---------------------------------------------------------------- playing */
static int save_wave(mdev *d, const char *file);
static void ensure_gp(mdev *d)
{
    if (d->gp || d->gp_failed || d->type == T_CD) return;
    if (d->tmpwav[0]) {                         /* recorded: played from a temporary file */
        if (save_wave(d, d->tmpwav)) { d->gp_failed = 1; return; }
        d->gp = gp_open(d->tmpwav, 0, 0);
        if (!d->gp) d->gp_failed = 1;
        return;
    }
    if (d->isnew) return;
    d->gp = gp_open(d->path, d->has_video, 0);
    if (!d->gp) d->gp_failed = 1;
    else if (d->audio_off) gp_volume(d->gp, 0);
}

static void stop_play(mdev *d, int code)
{
    if (!d->playing) return;
    d->pos = cur_pos(d);
    d->playing = 0;
    d->mode = M_STOP;
    gp_pause(d->gp);
    post(d, code);
}

/* devices that reached their end */
static void update(void)
{
    for (mdev *d = devs; d; d = d->next)
        if (d->playing && cur_pos(d) >= d->to) {
            d->pos = d->to; d->playing = 0; d->mode = M_STOP;
            gp_pause(d->gp);
            post(d, 1);
        }
}

static void wait_end(mdev *d)
{
    vb_files_flush();
    while (d->playing) {
        double left = nat_ms(d, d->to) - nat_ms(d, cur_pos(d));
        if (left > 0) usleep((useconds_t)(left > 20 ? 20000 : left * 1000 + 100));
        update();
    }
}

static void start_play(mdev *d, double from, double to)
{
    d->from = from; d->to = to; d->t0 = now_ms(); d->playing = 1; d->mode = M_PLAY; d->pos = from;
    if (d->type == T_CD) return;
    ensure_gp(d);
    if (d->gp) gp_play(d->gp, nat_ms(d, from), to >= d->len ? -1 : nat_ms(d, to));
}

/* ---------------------------------------------------------------- open */
static void file_type(const char *path, int *type)
{
    const char *e = strrchr(path, '.');
    if (!e) return;
    e++;
    if (!strcasecmp(e, "wav")) *type = T_WAVE;
    else if (!strcasecmp(e, "mid") || !strcasecmp(e, "rmi") || !strcasecmp(e, "midi")) *type = T_SEQ;
    else if (!strcasecmp(e, "avi")) *type = T_AVI;
    else if (!strcasecmp(e, "cda")) *type = T_CD;
    else {
        static const char *const mp[] = { "mp3", "mpg", "mpeg", "mp2", "mpa", "m1v", "mpe", "wma", "wmv", "asf", "au",
                                          "snd", "aif", "aiff", "aifc", "ogg", "flac", "mp4", "m4a", "mov", "mkv", "webm", 0 };
        for (int i = 0; mp[i]; i++) if (!strcasecmp(e, mp[i])) { *type = T_MPEG; return; }
    }
}
static int type_name(const char *s)
{
    for (int i = 1; i <= T_CD; i++) if (!strcasecmp(s, tnames[i])) return i;
    if (!strcasecmp(s, "digitalvideo")) return T_MPEG;
    return 0;
}

/* a path of the program (\\, any letter case, a drive letter) -> the file */
static void native(const char *in, char *out, size_t n)
{
    char p[1024];
    snprintf(p, sizeof p, "%s", in);
    for (char *q = p; *q; q++) if (*q == '\\') *q = '/';
    char *s = p;
    if (isalpha((unsigned char)s[0]) && s[1] == ':') s += 2;      /* C:/... */
    if (!access(s, F_OK)) { snprintf(out, n, "%s", s); return; }
    int abs = s[0] == '/';
    char res[2600], rest[1024];
    snprintf(res, sizeof res, "%s", abs ? "" : ".");
    snprintf(rest, sizeof rest, "%s", s);
    char *save = 0;
    for (char *part = strtok_r(rest, "/", &save); part; part = strtok_r(0, "/", &save)) {
        char cand[2600];
        snprintf(cand, sizeof cand, "%s/%s", res, part);
        if (access(cand, F_OK)) {
            DIR *dir = opendir(*res ? res : "/");
            struct dirent *e; int found = 0;
            while (dir && (e = readdir(dir)))
                if (!strcasecmp(e->d_name, part)) { snprintf(cand, sizeof cand, "%s/%s", res, e->d_name); found = 1; break; }
            if (dir) closedir(dir);
            if (!found) { snprintf(out, n, "%s", s); return; }
        }
        snprintf(res, sizeof res, "%s", cand);
    }
    snprintf(out, n, "%s", !abs && res[0] == '.' && res[1] == '/' ? res + 2 : res);
}

static int do_open(toks *k, char *ret, int retlen, mdev **pd)
{
    if (k->n < 2) return E_MISSING_DEVICE_NAME;
    char name[1024];
    snprintf(name, sizeof name, "%s", k->t[1]);
    k->used[0] = k->used[1] = 1;
    int type = 0, isnew = 0;
    const char *alias = 0;
    int share = 0;
    for (int i = 2; i < k->n; i++) {
        if (kw(k, i, "type") && i + 1 < k->n) {
            type = type_name(k->t[i + 1]);
            if (!type) return E_INVALID_DEVICE_NAME;
            k->used[i] = k->used[i + 1] = 1; i++;
        } else if (kw(k, i, "alias") && i + 1 < k->n) { alias = k->t[i + 1]; k->used[i] = k->used[i + 1] = 1; i++; }
        else if (kw(k, i, "shareable")) { share = 1; k->used[i] = 1; }
        else if (kw(k, i, "wait") || kw(k, i, "notify")) k->used[i] = 1;
        else if ((kw(k, i, "buffer") || kw(k, i, "parent") || kw(k, i, "style")) && i + 1 < k->n) { k->used[i] = k->used[i + 1] = 1; i++; }
        else if (kw(k, i, "type") || kw(k, i, "alias")) return E_MISSING_PARAMETER;
        else return E_UNRECOGNIZED_KEYWORD;
    }
    char elem[1024] = "";
    char *bang = strchr(name, '!');
    if (bang) { *bang = 0; type = type_name(name); if (!type) return E_INVALID_DEVICE_NAME; snprintf(elem, sizeof elem, "%s", bang + 1); }
    else if (!strcasecmp(name, "new")) { if (!type) return E_DEVICE_TYPE_REQUIRED; isnew = 1; }
    else if (!type && type_name(name)) type = type_name(name);    /* open cdaudio */
    else snprintf(elem, sizeof elem, "%s", name);
    if (isnew && !alias) return E_NEW_REQUIRES_ALIAS;
    if (!type && *elem) { file_type(elem, &type); if (!type) return E_EXTENSION_NOT_FOUND; }
    if (!type) return E_INVALID_DEVICE_NAME;
    const char *dname = alias ? alias : *elem ? elem : tnames[type];
    if (find_dev(dname)) return alias ? E_DUPLICATE_ALIAS : E_DEVICE_OPEN;
    if (isnew && type != T_WAVE) return E_UNSUPPORTED_FUNCTION;
    mdev *d = (mdev *)calloc(1, sizeof *d);
    d->type = type; d->isnew = isnew; d->shareable = share; d->speed = 1000;
    snprintf(d->alias, sizeof d->alias, "%s", dname);
    snprintf(d->elem, sizeof d->elem, "%s", elem);
    d->tf = type == T_AVI ? TF_FRAMES : type == T_CD ? TF_MSF : TF_MS;   /* sequencer: below */
    d->mode = M_STOP;
    int e = 0;
    if (*elem && type != T_CD) {
        native(elem, d->path, sizeof d->path);
        if (access(d->path, R_OK)) e = E_FILE_NOT_FOUND;
        else switch (type) {
        case T_WAVE: e = parse_wave(d); break;
        case T_SEQ: e = parse_midi(d); if (!d->smpte) d->tf = TF_SONGPTR; break;
        case T_AVI: e = parse_avi(d); break;
        case T_MPEG: {
            if (strrchr(d->path, '.') && !strcasecmp(strrchr(d->path, '.'), ".wav") && !parse_wave(d)) {
                d->len = d->len * 1000.0 / d->avgbps; d->has_audio = 1; d->avgbps = 0;
                break;
            }
            ensure_gp(d);
            double ms = gp_duration(d->gp);
            if (ms < 0) { e = d->gp ? E_INVALID_FILE : 0; ms = 0; }
            d->len = floor(ms);
            d->has_audio = 1;
            break;
        }
        }
    } else if (isnew) {
        d->ftag = 1; d->ch = 1; d->rate = 11025; d->bits = 8; d->align = 1; d->avgbps = 11025;
    } else if (type == T_CD) {
        cd_read_toc(d);
        d->tf = TF_MSF;
    } else if (type != T_CD) {
        e = E_FILENAME_REQUIRED;               /* waveaudio with no file */
        if (type == T_WAVE) { e = 0; d->mode = M_NOTREADY; }
    }
    if (e) { free(d->tmap); gp_close(d->gp); free(d); return e; }
    d->id = 1;                                  /* the lowest free number, as winmm */
    for (int again = 1; again;) { again = 0; for (mdev *x = devs; x; x = x->next) if (x->id == d->id) { d->id++; again = 1; } }
    mdev **pp = &devs;
    while (*pp) pp = &(*pp)->next;
    *pp = d;
    out_num(ret, retlen, d->id);
    *pd = d;
    return 0;
}

static void rec_stop(mdev *d);
static void free_dev(mdev *d)
{
    for (mdev **pp = &devs; *pp; pp = &(*pp)->next) if (*pp == d) { *pp = d->next; break; }
    rec_stop(d);
    if (d->tmpwav[0]) unlink(d->tmpwav);
    gp_close(d->gp);
    free(d->tmap); free(d->rec);
    free(d);
}

static void close_dev(mdev *d)
{
    if (d->playing) { d->playing = 0; post(d, 4); }
    free_dev(d);
}

/* ---------------------------------------------------------------- status */
static int do_status(mdev *d, toks *k, int i, char *r, int n)
{
    if (i >= k->n) return E_MISSING_PARAMETER;
    const char *w = k->t[i];
    int last = i;
#define IS(s) (!strcasecmp(w, s))
#define NEXT(s) (i + 1 < k->n && !strcasecmp(k->t[i + 1], s))
    if (IS("mode")) {
        const char *m = d->mode == M_NOTREADY ? "not ready" : d->playing ? "playing" : d->recording ? "recording" :
                        d->mode == M_PAUSE ? "paused" : d->mode == M_SEEK ? "seeking" : "stopped";
        if (d->type == T_CD && !d->ntrk) m = "open";
        snprintf(r, (size_t)n, "%s", m);
    } else if (IS("length")) {
        if (NEXT("track")) {
            if (i + 2 >= k->n) return E_MISSING_PARAMETER;
            int t = atoi(k->t[i + 2]); last = i + 2;
            if (d->type != T_CD) { if (t != 1) return E_OUTOFRANGE; fmt_len(d, d->len, r, n); }
            else {
                if (t < 1 || t > d->ntrk) return E_OUTOFRANGE;
                int save = d->tf; if (d->tf == TF_TMSF) d->tf = TF_MSF;
                long fr = d->trk[t - 1].len;
                if (d->tf == TF_MSF) snprintf(r, (size_t)n, "%02ld:%02ld:%02ld", fr / 4500, fr / 75 % 60, fr % 75);
                else fmt_len(d, fr, r, n);
                d->tf = save;
            }
        } else {
            if (d->type == T_CD && !d->ntrk) return E_DEVICE_NOT_READY;
            if (d->type == T_CD && (d->tf == TF_MSF || d->tf == TF_TMSF)) {
                long fr = (long)d->len;
                snprintf(r, (size_t)n, "%02ld:%02ld:%02ld", fr / 4500, fr / 75 % 60, fr % 75);
            } else { if (d->recording) rec_pull(d); fmt_len(d, d->len, r, n); }
        }
    } else if (IS("position")) {
        if (NEXT("track")) {
            if (i + 2 >= k->n) return E_MISSING_PARAMETER;
            int t = atoi(k->t[i + 2]); last = i + 2;
            if (d->type != T_CD) { if (t != 1) return E_OUTOFRANGE; fmt_time(d, 0, r, n); }
            else { if (t < 1 || t > d->ntrk) return E_OUTOFRANGE; fmt_time(d, d->trk[t - 1].start - d->trk[0].start, r, n); }
        } else if (d->type == T_SEQ && d->tf == TF_SONGPTR) out_num(r, n, 0);   /* MCISEQ gives 0 (VB6 oracle d_mci2) */
        else fmt_time(d, cur_pos(d), r, n);
    } else if (IS("start")) {
        if (!NEXT("position")) return E_MISSING_PARAMETER;
        last = i + 1;
        fmt_time(d, 0, r, n);
    } else if (IS("ready")) snprintf(r, (size_t)n, "%s", d->mode == M_NOTREADY || (d->type == T_CD && !d->ntrk) ? "false" : "true");
    else if (IS("media")) {
        if (!NEXT("present")) return E_MISSING_PARAMETER;
        last = i + 1;
        snprintf(r, (size_t)n, "%s", d->type == T_CD ? (d->ntrk ? "true" : "false") : "true");
    } else if (IS("number")) {
        if (!(NEXT("of") && i + 2 < k->n && !strcasecmp(k->t[i + 2], "tracks"))) return E_MISSING_PARAMETER;
        last = i + 2;
        out_num(r, n, d->type == T_CD ? d->ntrk : 1);
    } else if (IS("current")) {
        if (!NEXT("track")) return E_MISSING_PARAMETER;
        last = i + 1;
        int t = 1;
        if (d->type == T_CD) { double p = cur_pos(d) + (d->ntrk ? d->trk[0].start : 0); while (t < d->ntrk && d->trk[t].start <= p) t++; }
        out_num(r, n, t);
    } else if (IS("time")) {
        if (!NEXT("format")) return E_MISSING_PARAMETER;
        last = i + 1;
        snprintf(r, (size_t)n, "%s", tf_name(d->tf));
    } else if (d->type == T_WAVE && IS("channels")) out_num(r, n, d->ch);
    else if (d->type == T_WAVE && IS("samplespersec")) out_num(r, n, d->rate);
    else if (d->type == T_WAVE && IS("bitspersample")) out_num(r, n, d->bits);
    else if (d->type == T_WAVE && IS("bytespersec")) out_num(r, n, d->avgbps);
    else if (d->type == T_WAVE && IS("alignment")) out_num(r, n, d->align);
    else if (d->type == T_WAVE && IS("format")) {
        if (!NEXT("tag")) return E_MISSING_PARAMETER;
        last = i + 1;
        if (d->ftag == 1) snprintf(r, (size_t)n, "pcm"); else out_num(r, n, d->ftag);
    } else if (d->type == T_WAVE && IS("level")) out_num(r, n, 1);
    else if (d->type == T_WAVE && (IS("input") || IS("output"))) snprintf(r, (size_t)n, "%s", "mapper");
    else if (d->type == T_SEQ && IS("tempo")) out_num(r, n, d->smpte ? 0 : seq_tempo(d, cur_pos(d)));
    else if (d->type == T_SEQ && IS("port")) snprintf(r, (size_t)n, "%s", "mapper");
    else if (d->type == T_SEQ && (IS("master") || IS("slave"))) snprintf(r, (size_t)n, "%s", "none");
    else if (d->type == T_SEQ && IS("offset")) snprintf(r, (size_t)n, "%s", "00:00:00:00");
    else if ((d->type == T_AVI || d->type == T_MPEG) && IS("frame")) {
        if (!NEXT("rate")) return E_MISSING_PARAMETER;
        last = i + 1;
        out_num(r, n, d->usframe > 0 ? 1e9 / d->usframe : 0);
    } else if ((d->type == T_AVI || d->type == T_MPEG) && IS("nominal")) {
        if (!(NEXT("frame") && i + 2 < k->n && !strcasecmp(k->t[i + 2], "rate"))) return E_MISSING_PARAMETER;
        last = i + 2;
        out_num(r, n, d->usframe > 0 ? 1e9 / d->usframe : 0);
    } else if ((d->type == T_AVI || d->type == T_MPEG) && IS("speed")) out_num(r, n, d->speed);
    else if ((d->type == T_AVI || d->type == T_MPEG) && IS("window")) {
        if (NEXT("handle")) { last = i + 1; out_num(r, n, (double)d->hwnd); }
        else if (NEXT("visible")) { last = i + 1; snprintf(r, (size_t)n, "%s", d->winshow ? "true" : "false"); }
        else return E_MISSING_PARAMETER;
    } else if (d->type == T_CD && IS("type")) {
        if (!(NEXT("track") && i + 2 < k->n)) return E_MISSING_PARAMETER;
        int t = atoi(k->t[i + 2]); last = i + 2;
        if (t < 1 || t > d->ntrk) return E_OUTOFRANGE;
        snprintf(r, (size_t)n, "%s", d->trk[t - 1].audio ? "audio" : "other");
    } else return E_BAD_CONSTANT;
    for (int j = i; j <= last; j++) k->used[j] = 1;
    return 0;
#undef IS
#undef NEXT
}

static int do_capability(mdev *d, toks *k, int i, char *r, int n)
{
    if (i >= k->n) return E_MISSING_PARAMETER;
    const char *a = k->t[i], *b = i + 1 < k->n ? k->t[i + 1] : "";
    int last = i + 1, v = -1;
    int t = d->type, video = t == T_AVI || (t == T_MPEG && d->has_video);
    if (!strcasecmp(a, "can")) {
        if (!strcasecmp(b, "play")) v = 1;
        else if (!strcasecmp(b, "record")) v = t == T_WAVE;
        else if (!strcasecmp(b, "save")) v = t == T_WAVE;
        else if (!strcasecmp(b, "eject")) v = t == T_CD;
        else if (!strcasecmp(b, "stretch")) v = video;
        else if (!strcasecmp(b, "freeze")) v = 0;
        else if (!strcasecmp(b, "reverse")) v = 0;
        else if (!strcasecmp(b, "test")) v = video;
        else return E_BAD_CONSTANT;
    } else if (!strcasecmp(a, "has")) {
        if (!strcasecmp(b, "audio")) v = 1;
        else if (!strcasecmp(b, "video")) v = video;
        else return E_BAD_CONSTANT;
    } else if (!strcasecmp(a, "uses")) {
        if (!strcasecmp(b, "files")) v = t != T_CD;
        else if (!strcasecmp(b, "palettes")) v = video;
        else return E_BAD_CONSTANT;
    } else if (!strcasecmp(a, "compound")) {
        if (strcasecmp(b, "device")) return E_MISSING_PARAMETER;
        v = t != T_CD;
    } else if (!strcasecmp(a, "device")) {
        if (strcasecmp(b, "type")) return E_MISSING_PARAMETER;
        snprintf(r, (size_t)n, "%s", t == T_MPEG ? "digitalvideo" : t == T_AVI ? "digitalvideo" : tnames[t]);
        k->used[i] = k->used[i + 1] = 1;
        return 0;
    } else if (!strcasecmp(a, "inputs") || !strcasecmp(a, "outputs")) { out_num(r, n, 1); last = i; }
    else if (!strcasecmp(a, "windows")) { out_num(r, n, video ? 1 : 0); last = i; }
    else return E_BAD_CONSTANT;
    if (v >= 0) snprintf(r, (size_t)n, "%s", v ? "true" : "false");
    for (int j = i; j <= last && j < k->n; j++) k->used[j] = 1;
    return 0;
}

/* ---------------------------------------------------------------- set */
static int do_set1(mdev *d, toks *k, int i);
static int do_set(mdev *d, toks *k, int i)
{
    if (i >= k->n) return E_MISSING_PARAMETER;
    int e = do_set1(d, k, i);
    while (!e) {                                /* more items in the same command */
        int j = i; while (j < k->n && k->used[j]) j++;
        if (j >= k->n) break;
        if (j == i) break;
        i = j;
        e = do_set1(d, k, i);
    }
    return e;
}
static int do_set1(mdev *d, toks *k, int i)
{
    if (i >= k->n) return E_MISSING_PARAMETER;
    if (kw(k, i, "time") && kw(k, i + 1, "format")) {
        if (i + 2 >= k->n) return E_MISSING_PARAMETER;
        const char *f = k->t[i + 2]; int used = 3, tf = -1;
        if (!strcasecmp(f, "ms") || !strcasecmp(f, "milliseconds")) tf = TF_MS;
        else if (!strcasecmp(f, "bytes") && d->type == T_WAVE) tf = TF_BYTES;
        else if (!strcasecmp(f, "samples") && d->type == T_WAVE) tf = TF_SAMPLES;
        else if (!strcasecmp(f, "frames") && (d->type == T_AVI || (d->type == T_MPEG && d->has_video))) tf = TF_FRAMES;
        else if (!strcasecmp(f, "song") && kw(k, i + 3, "pointer") && d->type == T_SEQ) { tf = TF_SONGPTR; used = 4; }
        else if (!strcasecmp(f, "smpte") && i + 3 < k->n && d->type == T_SEQ) {
            used = 4;
            const char *x = k->t[i + 3];
            if (!strcmp(x, "24")) tf = TF_SMPTE24; else if (!strcmp(x, "25")) tf = TF_SMPTE25;
            else if (!strcmp(x, "30") && kw(k, i + 4, "drop")) { tf = TF_SMPTE30DROP; used = 5; }
            else if (!strcmp(x, "30")) tf = TF_SMPTE30;
        } else if (!strcasecmp(f, "msf") && d->type == T_CD) tf = TF_MSF;
        else if (!strcasecmp(f, "tmsf") && d->type == T_CD) tf = TF_TMSF;
        else if (!strcasecmp(f, "hms") && d->type == T_MPEG) tf = TF_HMS;
        if (tf < 0) return E_BAD_TIME_FORMAT;
        d->tf = tf;
        for (int j = i; j < i + used; j++) k->used[j] = 1;
        return 0;
    }
    if (kw(k, i, "audio")) {
        int j = i + 1;
        if (kw(k, j, "all") || kw(k, j, "left") || kw(k, j, "right")) j++;
        if (!kw(k, j, "on") && !kw(k, j, "off")) return E_MISSING_PARAMETER;
        d->audio_off = kw(k, j, "off");
        gp_volume(d->gp, d->audio_off ? 0 : 1);
        for (int q = i; q <= j; q++) k->used[q] = 1;
        return 0;
    }
    if (kw(k, i, "door")) {
        if (d->type != T_CD) return E_UNSUPPORTED_FUNCTION;
        if (!kw(k, i + 1, "open") && !kw(k, i + 1, "closed")) return E_MISSING_PARAMETER;
        k->used[i] = k->used[i + 1] = 1;
        return cd_door(kw(k, i, "open") || kw(k, i + 1, "open"));
    }
    if (kw(k, i, "video") && (kw(k, i + 1, "on") || kw(k, i + 1, "off"))) { k->used[i] = k->used[i + 1] = 1; return 0; }
    if (kw(k, i, "speed") && d->type != T_WAVE && d->type != T_CD) {
        if (i + 1 >= k->n) return E_MISSING_PARAMETER;
        d->speed = atof(k->t[i + 1]);
        k->used[i] = k->used[i + 1] = 1;
        return 0;
    }
    if (d->type == T_SEQ && kw(k, i, "tempo")) { if (i + 1 >= k->n) return E_MISSING_PARAMETER; k->used[i] = k->used[i + 1] = 1; return 0; }
    if (d->type == T_WAVE) {
        static const char *const wf[] = { "channels", "samplespersec", "bitspersample", "bytespersec", "alignment", 0 };
        for (int w = 0; wf[w]; w++)
            if (kw(k, i, wf[w])) {
                if (i + 1 >= k->n) return E_MISSING_PARAMETER;
                if (!d->isnew || d->len > 0) return E_NONAPPLICABLE_FUNCTION;
                int v = atoi(k->t[i + 1]);
                if (v <= 0) return E_OUTOFRANGE;
                switch (w) { case 0: d->ch = v; break; case 1: d->rate = v; break; case 2: d->bits = v; break;
                             case 3: d->avgbps = v; break; case 4: d->align = v; break; }
                if (w != 3 && w != 4) { d->align = d->ch * d->bits / 8; d->avgbps = d->rate * d->align; }
                k->used[i] = k->used[i + 1] = 1;
                return 0;
            }
        if (kw(k, i, "format") && kw(k, i + 1, "tag")) { k->used[i] = k->used[i + 1] = k->used[i + 2] = 1; return 0; }
    }
    return E_UNRECOGNIZED_KEYWORD;
}

/* ---------------------------------------------------------------- recording (waveaudio) */
static int save_wave(mdev *d, const char *file)
{
    char p[1024]; native(file, p, sizeof p);
    FILE *f = fopen(p, "wb");
    if (!f) return E_FILE_WRITE;
    unsigned len = (unsigned)d->nrec;
    unsigned char h[44];
    memcpy(h, "RIFF", 4);
    unsigned v = 36 + len;
#define W32(o, x) h[o] = (unsigned char)(x), h[o + 1] = (unsigned char)((x) >> 8), h[o + 2] = (unsigned char)((x) >> 16), h[o + 3] = (unsigned char)((x) >> 24)
#define W16(o, x) h[o] = (unsigned char)(x), h[o + 1] = (unsigned char)((x) >> 8)
    W32(4, v); memcpy(h + 8, "WAVEfmt ", 8); W32(16, 16u); W16(20, 1u); W16(22, (unsigned)d->ch); W32(24, (unsigned)d->rate);
    W32(28, (unsigned)d->avgbps); W16(32, (unsigned)d->align); W16(34, (unsigned)d->bits); memcpy(h + 36, "data", 4); W32(40, len);
#undef W32
#undef W16
    fwrite(h, 1, 44, f);
    if (len) fwrite(d->rec, 1, len, f);
    fclose(f);
    return 0;
}

/* ---------------------------------------------------------------- recording (waveaudio) */
static void rec_load(mdev *d)                   /* the samples of the file, once */
{
    if (d->rec || d->isnew || !*d->path) return;
    FILE *f = fopen(d->path, "rb");
    if (!f) return;
    unsigned char h[8];
    fseek(f, 12, SEEK_SET);
    while (fread(h, 1, 8, f) == 8) {
        unsigned sz = rd32(h + 4);
        if (!memcmp(h, "data", 4)) {
            size_t n = (size_t)d->len;
            d->rec = (unsigned char *)malloc(n + 1); d->caprec = n + 1;
            d->nrec = fread(d->rec, 1, n, f);
            break;
        }
        fseek(f, (long)(sz + (sz & 1)), SEEK_CUR);
    }
    fclose(f);
}
static void rec_put(mdev *d, const unsigned char *p, size_t n, int overwrite)
{
    if (d->recpos > d->nrec) d->recpos = d->nrec;
    size_t keep = overwrite ? (d->recpos + n < d->nrec ? d->nrec - d->recpos - n : 0) : d->nrec - d->recpos;
    size_t total = d->recpos + n + keep;
    if (total + 1 > d->caprec) { d->caprec = total * 2 + 4096; d->rec = (unsigned char *)realloc(d->rec, d->caprec); }
    memmove(d->rec + d->recpos + n, d->rec + d->nrec - keep, keep);
    if (p) memcpy(d->rec + d->recpos, p, n);
    else memset(d->rec + d->recpos, d->bits == 8 ? 0x80 : 0, n);    /* silence */
    d->recpos += n; d->nrec = total;
    d->len = (double)d->nrec;
}
static int rec_overwrite;
static void rec_pull(mdev *d)
{
    if (!d->recording || !d->rsink) return;
    void *smp;
    while ((smp = G.try_pull_sample(d->rsink, 0))) {
        void *buf = G.sample_get_buffer(smp);
        gmap m;
        if (buf && G.buffer_map(buf, &m, 1)) { rec_put(d, m.data, m.size, rec_overwrite); G.buffer_unmap(buf, &m); }
        G.mini_unref(smp);
    }
}
static void rec_start(mdev *d, int overwrite)
{
    rec_load(d);
    rec_overwrite = overwrite;
    d->recpos = (size_t)d->pos;
    d->recording = 1; d->mode = M_RECORD; d->t0 = now_ms();
    if (gst_load() && G.try_pull_sample && G.parse_launch) {
        const char *src = getenv("VBRT_MCI_SOURCE");
        char pl[600];
        snprintf(pl, sizeof pl, "%s ! audioconvert ! audioresample ! audio/x-raw,format=%s,rate=%d,channels=%d,layout=interleaved"
                 " ! appsink name=lbsink sync=false", src && *src ? src : "autoaudiosrc", d->bits == 8 ? "U8" : "S16LE", d->rate, d->ch);
        d->rpipe = G.parse_launch(pl, 0);
        if (d->rpipe) {
            d->rsink = G.bin_get_by_name(d->rpipe, "lbsink");
            G.set_state(d->rpipe, GST_PLAYING);
        }
    }
}
static void rec_stop(mdev *d)
{
    if (!d->recording) return;
    if (d->rpipe) {
        rec_pull(d);
        G.set_state(d->rpipe, GST_NULL);
        rec_pull(d);
        if (d->rsink) G.object_unref(d->rsink);
        G.object_unref(d->rpipe);
        d->rpipe = d->rsink = 0;
    } else {                                    /* no GStreamer: the time that passed, as silence */
        size_t n = (size_t)((now_ms() - d->t0) * d->avgbps / 1000.0);
        n -= n % (size_t)(d->align ? d->align : 1);
        rec_put(d, 0, n, rec_overwrite);
    }
    d->recording = 0; d->mode = M_STOP;
    d->pos = (double)d->recpos;
    /* to play what was recorded: a temporary file */
    if (!d->tmpwav[0]) snprintf(d->tmpwav, sizeof d->tmpwav, "%s/lbmci-%d-%d.wav", getenv("TMPDIR") ? getenv("TMPDIR") : "/tmp", (int)getpid(), d->id);
    gp_close(d->gp); d->gp = 0; d->gp_failed = 0;
}

/* ---------------------------------------------------------------- one command */
static vb_i32 run(char *cmd, char *ret, int retlen, vb_mci_notify cb, void *ctx, vb_i64 hwnd)
{
    toks k;
    if (ret && retlen > 0) ret[0] = 0;
    char dummy[4]; if (!ret || retlen <= 0) { ret = dummy; retlen = sizeof dummy; }
    int e = tokenize(cmd, &k);
    if (e) return e;
    if (k.n == 0) return E_MISSING_COMMAND_STRING;
    const char *verb = k.t[0];
    int notify = find_kw(&k, 1, "notify") >= 0, wait = find_kw(&k, 1, "wait") >= 0;
    update();
    if (!strcasecmp(verb, "open")) {
        mdev *d = 0;
        e = do_open(&k, ret, retlen, &d);
        if (!e && notify) { d->cb = cb; d->ctx = ctx; d->nhwnd = hwnd; d->npending = 1; post(d, 1); }
        return e;
    }
    if (!strcasecmp(verb, "sysinfo")) {
        if (k.n < 3) return E_MISSING_PARAMETER;
        int open_only = find_kw(&k, 3, "open") >= 0;
        int ty = strcasecmp(k.t[1], "all") ? type_name(k.t[1]) : -1;
        if (!strcasecmp(k.t[2], "quantity")) {
            int c = 0;
            if (open_only) { for (mdev *d = devs; d; d = d->next) if (ty < 0 || d->type == ty) c++; }
            else c = ty < 0 ? 5 : ty ? 1 : 0;
            out_num(ret, retlen, c);
            return 0;
        }
        if (!strcasecmp(k.t[2], "name") && k.n >= 4) {
            int want = atoi(k.t[3]), c = 0;
            if (open_only) {
                for (mdev *d = devs; d; d = d->next) if ((ty < 0 || d->type == ty) && ++c == want) { snprintf(ret, (size_t)retlen, "%s", d->alias); return 0; }
            } else if (ty < 0 ? (want >= 1 && want <= 5) : want == 1) { snprintf(ret, (size_t)retlen, "%s", tnames[ty < 0 ? want : ty]); return 0; }
            return E_OUTOFRANGE;
        }
        if (!strcasecmp(k.t[2], "installname")) {
            mdev *d = find_dev(k.t[1]);
            if (d) { snprintf(ret, (size_t)retlen, "%s", tnames[d->type]); return 0; }
            if (ty > 0) { snprintf(ret, (size_t)retlen, "%s", tnames[ty]); return 0; }
            return E_INVALID_DEVICE_NAME;
        }
        return E_UNRECOGNIZED_KEYWORD;
    }
    if (!strcasecmp(verb, "close") && k.n >= 2 && !strcasecmp(k.t[1], "all")) {
        while (devs) close_dev(devs);
        return 0;
    }
    static const char *const verbs[] = { "close", "play", "stop", "pause", "resume", "seek", "status", "set", "capability",
        "info", "record", "save", "window", "where", "put", "step", "cue", "delete", "load", "setaudio", "setvideo",
        "signal", "update", "realize", "configure", "freeze", "unfreeze", "break", "sound", "spin", "escape", "list",
        "monitor", "settimecode", "settuner", "undo", "redo", "copy", "cut", "paste", "capture", "reserve", "restore", 0 };
    int known = 0;
    for (int i = 0; verbs[i]; i++) if (!strcasecmp(verb, verbs[i])) known = 1;
    if (!known) return E_UNRECOGNIZED_COMMAND;
    if (k.n < 2) return E_MISSING_DEVICE_NAME;
    mdev *d = find_dev(k.t[1]);
    int autoopened = 0;
    if (!d) {
        /* auto-open: "play c:\x.wav" */
        if (!strcasecmp(verb, "play") || !strcasecmp(verb, "status") || !strcasecmp(verb, "capability") || !strcasecmp(verb, "info")) {
            int ty = 0; file_type(k.t[1], &ty);
            char p[1024]; native(k.t[1], p, sizeof p);
            if (ty && !access(p, R_OK) && !strcasecmp(verb, "play")) {
                if (notify) return E_NOTIFY_ON_AUTO_OPEN;
                char oc[1200]; snprintf(oc, sizeof oc, "open \"%s\"", k.t[1]);
                toks k2; char r2[32]; mdev *nd = 0;
                if (tokenize(oc, &k2) || do_open(&k2, r2, sizeof r2, &nd)) return E_INVALID_DEVICE_NAME;
                d = nd; d->autoopen = 1; autoopened = 1;
            }
        }
        if (!d) return E_INVALID_DEVICE_NAME;
    }
    k.used[0] = k.used[1] = 1;
    int nf = find_kw(&k, 2, "notify"); if (nf >= 0) k.used[nf] = 1;
    int wf = find_kw(&k, 2, "wait"); if (wf >= 0) k.used[wf] = 1;
    if (notify && d->npending)                  /* superseded; a Stop aborts a Play */
        post(d, !strcasecmp(verb, "stop") && d->playing ? 4 : 2);
    e = 0;
    int keep_pending = 0;
    if (!strcasecmp(verb, "close")) {
        if (notify) { d->cb = cb; d->ctx = ctx; d->nhwnd = hwnd; d->npending = 1; post(d, 1); }
        close_dev(d);
        return 0;
    } else if (!strcasecmp(verb, "play")) {
        if (d->mode == M_NOTREADY || (d->isnew && d->len <= 0)) e = d->isnew ? 0 : E_DEVICE_NOT_READY;
        if (!e && d->type == T_CD && !d->ntrk) e = E_DEVICE_NOT_READY;
        double from = cur_pos(d), to = d->len;
        int fi = find_kw(&k, 2, "from"), ti = find_kw(&k, 2, "to");
        if (!e && fi >= 0) { if (fi + 1 >= k.n) e = E_MISSING_PARAMETER; else { e = parse_time(d, k.t[fi + 1], &from); k.used[fi] = k.used[fi + 1] = 1; } }
        if (!e && ti >= 0) { if (ti + 1 >= k.n) e = E_MISSING_PARAMETER; else { e = parse_time(d, k.t[ti + 1], &to); k.used[ti] = k.used[ti + 1] = 1; } }
        if (!e) for (int i = 2; i < k.n; i++) if (!k.used[i] && (kw(&k, i, "fullscreen") || kw(&k, i, "window") || kw(&k, i, "repeat") || kw(&k, i, "reverse"))) k.used[i] = 1;
        if (!e && (from > d->len || to > d->len)) e = E_OUTOFRANGE;
        if (!e && to < from && ti >= 0) e = E_OUTOFRANGE;
        if (!e && fi < 0 && from >= d->len && ti < 0) from = d->len;
        if (!e) {
            if (d->playing) { d->playing = 0; if (d->npending && notify == 0) post(d, 4); }
            if (notify) { d->cb = cb; d->ctx = ctx; d->nhwnd = hwnd; }
            if (to <= from && d->len > 0) { d->pos = from; d->mode = M_STOP; if (notify) { d->npending = 1; post(d, 1); } }
            else {
                start_play(d, from, to);
                if (notify) { d->npending = 1; keep_pending = 1; }
                if (wait || autoopened) wait_end(d);
            }
            if (autoopened && !d->playing) { free_dev(d); return 0; }
        }
    } else if (!strcasecmp(verb, "stop")) {
        if (d->playing) stop_play(d, 4);
        else if (d->npending) post(d, 4);
        if (d->mode == M_PAUSE) d->mode = M_STOP;
        if (d->recording) rec_stop(d);
    } else if (!strcasecmp(verb, "pause")) {
        if (d->playing) { d->pos = cur_pos(d); d->playing = 0; d->mode = M_PAUSE; gp_pause(d->gp); if (d->npending) keep_pending = 1; }
        else if (d->mode != M_PAUSE) e = E_NONAPPLICABLE_FUNCTION;
    } else if (!strcasecmp(verb, "resume")) {
        if (d->mode == M_PAUSE) { start_play(d, d->pos, d->to > d->pos ? d->to : d->len); if (d->npending) keep_pending = 1; }
        else if (!d->playing) e = E_NONAPPLICABLE_FUNCTION;
    } else if (!strcasecmp(verb, "seek")) {
        double to = -1;
        int ti = find_kw(&k, 2, "to");
        if (ti < 0) e = E_MISSING_PARAMETER;
        else if (ti + 1 >= k.n) e = E_MISSING_PARAMETER;
        else if (kw(&k, ti + 1, "start")) to = 0;
        else if (kw(&k, ti + 1, "end")) to = d->len;
        else e = parse_time(d, k.t[ti + 1], &to);
        if (!e) { k.used[ti] = k.used[ti + 1] = 1; if (to > d->len) e = E_OUTOFRANGE; }
        if (!e) {
            if (d->playing) stop_play(d, 4);
            /* MCISEQ: a seek in song pointers of a format 1 file goes to the start (VB6 oracle d_mci2) */
            if (d->type == T_SEQ && d->tf == TF_SONGPTR && d->midifmt == 1 && ti + 1 < k.n && !kw(&k, ti + 1, "end")) to = 0;
            d->pos = to; d->mode = M_STOP; d->tempo_moved = 1;
            gp_seek(d->gp, nat_ms(d, to));
        }
    } else if (!strcasecmp(verb, "status")) {
        e = d->autoopen ? E_ILLEGAL_FOR_AUTO_OPEN : 0;
        if (!e) {
            int i = 2; while (i < k.n && k.used[i]) i++;
            if (i >= k.n) e = E_MISSING_PARAMETER;
            else e = do_status(d, &k, i, ret, retlen);
        }
    } else if (!strcasecmp(verb, "capability")) {
        int i = 2; while (i < k.n && k.used[i]) i++;
        e = do_capability(d, &k, i, ret, retlen);
    } else if (!strcasecmp(verb, "info")) {
        int i = 2; while (i < k.n && k.used[i]) i++;
        if (i >= k.n) e = E_MISSING_PARAMETER;
        else if (kw(&k, i, "file")) { if (d->type == T_CD) e = E_UNSUPPORTED_FUNCTION; else snprintf(ret, (size_t)retlen, "%s", d->elem); }
        else if (kw(&k, i, "product")) snprintf(ret, (size_t)retlen, "%s",
            d->type == T_WAVE ? "WAVE Audio" : d->type == T_SEQ ? "MIDI Sequencer" : d->type == T_AVI ? "Video for Windows" :
            d->type == T_CD ? "CD Audio" : "DirectShow");
        else if (kw(&k, i, "version")) snprintf(ret, (size_t)retlen, "%s", "6.1");
        else if (kw(&k, i, "window") && kw(&k, i + 1, "text")) { if (d->type == T_AVI || d->type == T_MPEG) { snprintf(ret, (size_t)retlen, "%s", d->alias); k.used[i + 1] = 1; } else e = E_UNSUPPORTED_FUNCTION; }
        else if (kw(&k, i, "identity") && d->type == T_CD) snprintf(ret, (size_t)retlen, "%08X", d->ntrk ? (unsigned)(d->len * 13 + d->ntrk) : 0u);
        else e = E_UNRECOGNIZED_KEYWORD;
        if (!e) k.used[i] = 1;
    } else if (!strcasecmp(verb, "set")) {
        int i = 2; while (i < k.n && k.used[i]) i++;
        e = do_set(d, &k, i);
    } else if (!strcasecmp(verb, "record")) {
        if (d->type != T_WAVE) e = E_UNSUPPORTED_FUNCTION;
        else {
            for (int i = 2; i < k.n; i++) if (!k.used[i] && (kw(&k, i, "insert") || kw(&k, i, "overwrite"))) k.used[i] = 1;
            for (int i = 2; i + 1 < k.n; i++) if (!k.used[i] && (kw(&k, i, "from") || kw(&k, i, "to"))) k.used[i] = k.used[i + 1] = 1;
            if (d->playing) stop_play(d, 4);
            rec_start(d, find_kw(&k, 2, "overwrite") >= 0);
        }
    } else if (!strcasecmp(verb, "save")) {
        if (d->type != T_WAVE) e = E_UNSUPPORTED_FUNCTION;
        else {
            if (d->recording) rec_stop(d);
            rec_load(d);
            int i = 2; while (i < k.n && k.used[i]) i++;
            const char *f = i < k.n ? k.t[i] : d->elem;
            if (i < k.n) k.used[i] = 1;
            if (!*f) e = E_FILENAME_REQUIRED;
            else e = save_wave(d, f);
        }
    } else if (!strcasecmp(verb, "window")) {
        if (d->type != T_AVI && d->type != T_MPEG) e = E_UNSUPPORTED_FUNCTION;
        else {
            for (int i = 2; i < k.n; i++) {
                if (k.used[i]) continue;
                if (kw(&k, i, "handle") && i + 1 < k.n) {
                    if (kw(&k, i + 1, "default")) d->hwnd = 0; else d->hwnd = strtoll(k.t[i + 1], 0, 10);
                    k.used[i] = k.used[i + 1] = 1; i++;
                } else if (kw(&k, i, "state") && i + 1 < k.n) {
                    d->winshow = !kw(&k, i + 1, "hide") && !kw(&k, i + 1, "minimize");
                    k.used[i] = k.used[i + 1] = 1; i++;
                } else if (kw(&k, i, "text") && i + 1 < k.n) { k.used[i] = k.used[i + 1] = 1; i++; }
            }
        }
    } else if (!strcasecmp(verb, "where")) {
        if (d->type != T_AVI && d->type != T_MPEG) e = E_UNSUPPORTED_FUNCTION;
        else {
            int i = 2; while (i < k.n && k.used[i]) i++;
            if (i < k.n && (kw(&k, i, "source") || kw(&k, i, "destination") || kw(&k, i, "window"))) {
                /* the destination of the default window: its client area (Windows 7: 116 pixels at least) */
                int dw = kw(&k, i, "source") || d->hwnd ? d->vw : (d->vw < 116 ? 116 : d->vw);
                snprintf(ret, (size_t)retlen, "0 0 %d %d", dw, d->vh);
                k.used[i] = 1;
            } else e = E_MISSING_PARAMETER;
        }
    } else if (!strcasecmp(verb, "put")) {
        if (d->type != T_AVI && d->type != T_MPEG) e = E_UNSUPPORTED_FUNCTION;
        else for (int i = 2; i < k.n; i++) k.used[i] = 1;
    } else if (!strcasecmp(verb, "step")) {
        if (d->type != T_AVI && d->type != T_MPEG) e = E_UNSUPPORTED_FUNCTION;
        else {
            int by = 1, rev = find_kw(&k, 2, "reverse");
            int bi = find_kw(&k, 2, "by");
            if (bi >= 0 && bi + 1 < k.n) { by = atoi(k.t[bi + 1]); k.used[bi] = k.used[bi + 1] = 1; }
            if (rev >= 0) { k.used[rev] = 1; by = -by; }
            if (d->playing) stop_play(d, 4);
            double p = d->pos + by;
            if (p < 0) p = 0;
            if (p > d->len) p = d->len;
            d->pos = p;
            gp_seek(d->gp, nat_ms(d, p));
        }
    } else if (!strcasecmp(verb, "cue") || !strcasecmp(verb, "realize") || !strcasecmp(verb, "update") ||
               !strcasecmp(verb, "signal") || !strcasecmp(verb, "setaudio") || !strcasecmp(verb, "setvideo")) {
        for (int i = 2; i < k.n; i++) k.used[i] = 1;
        if (d->type == T_CD && strcasecmp(verb, "cue")) e = E_UNSUPPORTED_FUNCTION;
    } else e = E_UNSUPPORTED_FUNCTION;
    if (!e) for (int i = 2; i < k.n; i++) if (!k.used[i]) { e = E_UNRECOGNIZED_KEYWORD; break; }
    if (!e && notify && !keep_pending) { d->cb = cb; d->ctx = ctx; d->nhwnd = hwnd; d->npending = 1; post(d, 1); }
    if (e && notify && d->npending && !keep_pending) { /* a failed command does not notify */ }
    return e;
}

vb_i32 vb_mci_send_hwnd(const char *cmd, char *ret, vb_i32 retlen, vb_mci_notify cb, void *ctx, vb_i64 hwnd)
{
    char *c = strdup(cmd ? cmd : "");
    vb_i32 r = run(c, ret, retlen, cb, ctx, hwnd);
    free(c);
    return r;
}
vb_i32 vb_mci_send(const char *cmd, char *ret, vb_i32 retlen, vb_mci_notify cb, void *ctx)
{
    return vb_mci_send_hwnd(cmd, ret, retlen, cb, ctx, 0);
}

/* the notifications that are due: called by the forms runtime (a timer) and by DoEvents */
void vb_mci_poll(void)
{
    update();
    for (mdev *d = devs; d; d = d->next) if (d->gp) gp_drain(d->gp);
    while (nqueue) {
        nq q = queue[0];
        memmove(queue, queue + 1, sizeof(nq) * (size_t)--nqueue);
        if (q.cb) q.cb(q.ctx, q.id, q.code);
        else if (q.hwnd && vb_mci_post_hook) vb_mci_post_hook(q.hwnd, q.code, q.id);
    }
}
void vb_mci_forget(void *ctx)
{
    for (int i = 0; i < nqueue; i++) if (queue[i].ctx == ctx) queue[i].cb = 0, queue[i].hwnd = 0;
    for (mdev *d = devs; d; d = d->next) if (d->ctx == ctx) { d->cb = 0; d->ctx = 0; d->npending = 0; }
}
vb_i32 vb_mci_active(void)
{
    if (nqueue) return 1;
    for (mdev *d = devs; d; d = d->next) if (d->playing || d->npending) return 1;
    return 0;
}

/* video: the devices with pictures, and the newest picture of one */
vb_i32 vb_mci_next_video(vb_i32 after)
{
    int found = after == 0;
    for (mdev *d = devs; d; d = d->next) {
        if (found && d->has_video) return d->id;
        if (d->id == after) found = 1;
    }
    return 0;
}
int vb_mci_frame(vb_i32 id, vb_mci_frameinfo *f)
{
    mdev *d = 0;
    for (mdev *x = devs; x; x = x->next) if (x->id == id) d = x;
    if (!d) return 0;
    f->hwnd = d->hwnd; f->show = d->winshow || d->playing; f->w = d->vw; f->h = d->vh; f->data = 0; f->stride = 0;
    f->playing = d->playing;
    snprintf(f->title, sizeof f->title, "%s", d->alias);
    ensure_gp(d);
    if (!d->gp || !d->gp->vsink) return 1;
    void *s = G.try_pull_sample(d->gp->vsink, 0);
    if (!s && G.try_pull_preroll) s = G.try_pull_preroll(d->gp->vsink, 0);
    if (!s) return 1;
    void *caps = G.sample_get_caps(s), *buf = G.sample_get_buffer(s);
    int w = 0, h = 0;
    if (caps) { void *st = G.caps_get_structure(caps, 0); G.structure_get_int(st, "width", &w); G.structure_get_int(st, "height", &h); }
    static unsigned char *keep; static size_t nkeep;
    gmap m;
    if (buf && w > 0 && h > 0 && G.buffer_map(buf, &m, 1)) {
        if (nkeep < m.size) { keep = (unsigned char *)realloc(keep, m.size); nkeep = m.size; }
        memcpy(keep, m.data, m.size);
        G.buffer_unmap(buf, &m);
        f->w = w; f->h = h; f->data = keep; f->stride = (int)(m.size / (size_t)h); f->fresh = 1;
    }
    G.mini_unref(s);
    return 1;
}
#endif

/* ---------------------------------------------------------------- Declare: winmm */
vb_i32 vb_mci_device_id(const char *name)
{
#ifdef _WIN32
    return lbmci_devid(name);
#else
    update();
    mdev *d = find_dev(name);
    return d ? d->id : 0;
#endif
}
