/* lbmapi_dll.c - LinBASIC's provider of Simple MAPI (runtime/lbmapi.c) as a DLL of Windows: the mail program of the
 * VB6 tests of phase 7.9 in the VM (install.bat makes it the default mail program, uninstall.bat takes it away).
 * The functions of MAPI.H with their own names (built with -Wl,--kill-at: MAPILogon, not _MAPILogon@24).
 * Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later. */
#include <windows.h>
#include "../../runtime/lbmapi.h"

#define MAPIFN __declspec(dllexport) ULONG WINAPI

MAPIFN MAPILogon(ULONG_PTR ui, LPSTR profile, LPSTR password, ULONG flags, ULONG reserved, ULONG_PTR *session)
{
    (void)reserved;
    lbm_handle h = 0;
    ULONG r = lbm_logon(ui, profile, password, flags, &h);
    if (session) *session = h;
    return r;
}
MAPIFN MAPILogoff(ULONG_PTR session, ULONG_PTR ui, ULONG flags, ULONG reserved)
{
    (void)reserved;
    return lbm_logoff(session, ui, flags);
}
MAPIFN MAPISendMail(ULONG_PTR session, ULONG_PTR ui, lbm_message *msg, ULONG flags, ULONG reserved)
{
    (void)reserved;
    return lbm_send_mail(session, ui, msg, flags);
}
MAPIFN MAPISendDocuments(ULONG_PTR ui, LPSTR delims, LPSTR paths, LPSTR names, ULONG reserved)
{
    (void)reserved;
    return lbm_send_documents(ui, delims, paths, names);
}
MAPIFN MAPIFindNext(ULONG_PTR session, ULONG_PTR ui, LPSTR type, LPSTR seed, ULONG flags, ULONG reserved, LPSTR id)
{
    (void)reserved;
    char buf[512];
    ULONG r = lbm_find_next(session, ui, type, seed, flags, buf);
    if (id) lstrcpynA(id, r ? "" : buf, (flags & LBM_LONG_MSGID) ? 512 : 64);
    return r;
}
MAPIFN MAPIReadMail(ULONG_PTR session, ULONG_PTR ui, LPSTR id, ULONG flags, ULONG reserved, lbm_message **msg)
{
    (void)reserved;
    return lbm_read_mail(session, ui, id, flags, msg);
}
MAPIFN MAPISaveMail(ULONG_PTR session, ULONG_PTR ui, lbm_message *msg, ULONG flags, ULONG reserved, LPSTR id)
{
    (void)reserved;
    char buf[512];
    lstrcpynA(buf, id ? id : "", sizeof buf);
    ULONG r = lbm_save_mail(session, ui, msg, flags, buf);
    if (id && !r) lstrcpynA(id, buf, (flags & LBM_LONG_MSGID) ? 512 : 64);
    return r;
}
MAPIFN MAPIDeleteMail(ULONG_PTR session, ULONG_PTR ui, LPSTR id, ULONG flags, ULONG reserved)
{
    (void)reserved;
    return lbm_delete_mail(session, ui, id, flags);
}
MAPIFN MAPIFreeBuffer(LPVOID p)
{
    return lbm_free_buffer(p);
}
MAPIFN MAPIAddress(ULONG_PTR session, ULONG_PTR ui, LPSTR caption, ULONG nedit, LPSTR labels, ULONG nrecips,
                   lbm_recip *recips, ULONG flags, ULONG reserved, ULONG *nnew, lbm_recip **news)
{
    (void)reserved;
    lbm_ulong n = 0;
    ULONG r = lbm_address(session, ui, caption, nedit, labels, nrecips, recips, flags, &n, news);
    if (nnew) *nnew = n;
    return r;
}
MAPIFN MAPIDetails(ULONG_PTR session, ULONG_PTR ui, lbm_recip *recip, ULONG flags, ULONG reserved)
{
    (void)reserved;
    return lbm_details(session, ui, recip, flags);
}
MAPIFN MAPIResolveName(ULONG_PTR session, ULONG_PTR ui, LPSTR name, ULONG flags, ULONG reserved, lbm_recip **recip)
{
    (void)reserved;
    return lbm_resolve_name(session, ui, name, flags, recip);
}
BOOL WINAPI DllMain(HINSTANCE h, DWORD why, LPVOID r) { (void)h; (void)why; (void)r; return TRUE; }
