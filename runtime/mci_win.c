/* mci_win.c - MCI on Windows: the real winmm, for mci.c (the MMControl and the MCI runtime);
 * apart from mci.c because windows.h and vbrt.h both name VT_EMPTY ...
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifdef _WIN32
#include <windows.h>
#include <mmsystem.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>

typedef void (*ncb)(void *, int, int);
typedef struct { ncb cb; void *ctx; } wnot;
static HWND notify_wnd;
static wnot *wn; static int nwn;              /* by device id */

static LRESULT CALLBACK nproc(HWND h, UINT m, WPARAM w, LPARAM l)
{
    if (m == MM_MCINOTIFY) {
        int id = (int)l;
        if (id > 0 && id < nwn && wn[id].cb) wn[id].cb(wn[id].ctx, id, (int)w);
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

int lbmci_send(const char *cmd, char *ret, int retlen, ncb cb, void *ctx, long long hwnd)
{
    HWND cbw = (HWND)(intptr_t)hwnd;
    if (cb) {
        if (!notify_wnd) {
            WNDCLASSW c; memset(&c, 0, sizeof c);
            c.lpfnWndProc = nproc; c.hInstance = GetModuleHandleW(0); c.lpszClassName = L"LinBasicMci";
            RegisterClassW(&c);
            notify_wnd = CreateWindowW(L"LinBasicMci", L"", 0, 0, 0, 0, 0, HWND_MESSAGE, 0, c.hInstance, 0);
        }
        cbw = notify_wnd;
    }
    int n = MultiByteToWideChar(CP_UTF8, 0, cmd, -1, 0, 0);
    wchar_t *w = (wchar_t *)malloc(sizeof(wchar_t) * (n + 1));
    MultiByteToWideChar(CP_UTF8, 0, cmd, -1, w, n);
    wchar_t r[512]; r[0] = 0;
    MCIERROR e = mciSendStringW(w, r, 512, cbw);
    free(w);
    if (ret && retlen > 0) { ret[0] = 0; WideCharToMultiByte(CP_UTF8, 0, r, -1, ret, retlen, 0, 0); ret[retlen - 1] = 0; }
    if (cb && !e) {                             /* the device of the command: its notifications go to cb */
        wchar_t dev[256];
        char a[256]; int k = 0; const char *q = cmd;
        while (*q == ' ') q++;
        while (*q && *q != ' ') q++;           /* the verb */
        while (*q == ' ') q++;
        if (*q == '"') { q++; while (*q && *q != '"' && k < 255) a[k++] = *q++; }
        else while (*q && *q != ' ' && k < 255) a[k++] = *q++;
        a[k] = 0;
        MultiByteToWideChar(CP_UTF8, 0, a, -1, dev, 256);
        int id = (int)mciGetDeviceIDW(dev);
        if (!id) {                              /* open ... alias x */
            const char *al = strstr(cmd, " alias ");
            if (al) { al += 7; k = 0; while (*al == ' ') al++; while (*al && *al != ' ' && k < 255) a[k++] = *al++; a[k] = 0;
                      MultiByteToWideChar(CP_UTF8, 0, a, -1, dev, 256); id = (int)mciGetDeviceIDW(dev); }
        }
        if (id > 0) {
            if (id >= nwn) { int nn = id + 16; wn = (wnot *)realloc(wn, sizeof(wnot) * nn); memset(wn + nwn, 0, sizeof(wnot) * (nn - nwn)); nwn = nn; }
            wn[id].cb = cb; wn[id].ctx = ctx;
        }
    }
    return (int)e;
}
int lbmci_errtext(int code, char *buf, int len)
{
    wchar_t w[512];
    if (!mciGetErrorStringW((MCIERROR)code, w, 512)) { if (buf && len) buf[0] = 0; return 0; }
    WideCharToMultiByte(CP_UTF8, 0, w, -1, buf, len, 0, 0);
    return 1;
}
void lbmci_forget(void *ctx)
{
    for (int i = 0; i < nwn; i++) if (wn[i].ctx == ctx) wn[i].cb = 0;
}
int lbmci_devid(const char *name)
{
    wchar_t w[512];
    MultiByteToWideChar(CP_UTF8, 0, name, -1, w, 512);
    return (int)mciGetDeviceIDW(w);
}
#endif
