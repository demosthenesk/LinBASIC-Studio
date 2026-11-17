/* launcher.c - "LinBASIC Studio.exe" of the portable package of Windows: starts the IDE with the
 * Python of the package (python\pythonw.exe -m lbstudio <arguments>), from the folder of the .exe.
 * Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
 *   i686-w64-mingw32-gcc -municode -mwindows -Os -s launcher.c -o "LinBASIC Studio.exe" */
#include <windows.h>
#include <wchar.h>

int WINAPI wWinMain(HINSTANCE inst, HINSTANCE prev, PWSTR args, int show)
{
    static wchar_t dir[32768], line[65536];
    (void)inst; (void)prev; (void)show;
    DWORD n = GetModuleFileNameW(NULL, dir, 32768);
    if (n == 0 || n >= 32768) return 1;
    wchar_t *slash = wcsrchr(dir, L'\\');
    if (slash) *slash = 0;
    _snwprintf(line, 65535, L"\"%ls\\python\\pythonw.exe\" -m lbstudio %ls", dir, args ? args : L"");
    line[65535] = 0;
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof si);
    si.cb = sizeof si;
    if (!CreateProcessW(NULL, line, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        wchar_t msg[1024];
        _snwprintf(msg, 1023, L"Can't start %ls\\python\\pythonw.exe (error %lu).\nThe folder of LinBASIC Studio "
                   L"is incomplete: unpack the whole zip again.", dir, GetLastError());
        msg[1023] = 0;
        MessageBoxW(NULL, msg, L"LinBASIC Studio", MB_ICONERROR);
        return 1;
    }
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return 0;
}
