/* dbgpipe_win.c - the connection of debug.c to the IDE on Windows: the named pipe of QLocalServer
 * (\\.\pipe\<name>). Apart from debug.c because windows.h and vbrt.h both name VT_EMPTY ...
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <windows.h>
#include <stddef.h>

static HANDLE pipe_h = INVALID_HANDLE_VALUE;
int vb_dbgpipe_open(const char *path)
{
    for (int i = 0; i < 50; i++) {
        pipe_h = CreateFileA(path, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
        if (pipe_h != INVALID_HANDLE_VALUE) return 0;
        if (GetLastError() != ERROR_PIPE_BUSY) return -1;
        WaitNamedPipeA(path, 200);
    }
    return -1;
}
long vb_dbgpipe_write(const char *s, size_t n)
{
    DWORD k = 0;
    if (!WriteFile(pipe_h, s, (DWORD)n, &k, NULL)) return -1;
    return (long)k;
}
long vb_dbgpipe_read(char *s, size_t n)
{
    DWORD k = 0;
    if (!ReadFile(pipe_h, s, (DWORD)n, &k, NULL)) return -1;
    return (long)k;
}
/* data to read within timeout_ms (< 0: wait) */
int vb_dbgpipe_wait(int timeout_ms)
{
    DWORD start = GetTickCount();
    for (;;) {
        DWORD avail = 0;
        if (!PeekNamedPipe(pipe_h, NULL, 0, NULL, &avail, NULL)) return 1;   /* closed: read says so */
        if (avail) return 1;
        if (timeout_ms == 0 || (timeout_ms > 0 && GetTickCount() - start >= (DWORD)timeout_ms)) return 0;
        Sleep(5);
    }
}
void vb_dbgpipe_close(void)
{
    if (pipe_h != INVALID_HANDLE_VALUE) CloseHandle(pipe_h);
    pipe_h = INVALID_HANDLE_VALUE;
}
