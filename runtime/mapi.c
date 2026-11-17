/* mapi.c - the Simple MAPI that the MAPI controls of the forms runtime (ocx_mapi.inc: MAPISession, MAPIMessages)
 * call, phase 7.9. Where the calls go:
 *   - Linux: LinBASIC's own provider (lbmapi.c: the messages of a store folder, sending by sendmail / SMTP /
 *     xdg-email);
 *   - Windows: the MAPI32.DLL of the system (the mail program the user chose), as VB6's MSMAPI32.OCX does; lbmapi
 *     when the environment variable LBMAPI_DIR names a store (the tests in Wine);
 *   - LBMAPI_NONE=1 (the tests of a computer without a mail program): no provider; every call fails as the MAPI32
 *     stub of Windows does without one (0x4005, the control's error 48389).
 * The functions are those of lbmapi.h, with lbmc_ for lbm_.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "lbmapi.h"
#include <stdlib.h>
#include <string.h>
#ifdef _WIN32
#include <windows.h>
#endif

#define NO_PROVIDER 0x4005                      /* E_FAIL of the stub (its low word) */

enum { M_UNSET, M_NONE, M_OWN, M_MAPI32 };
static int mode = M_UNSET;

#ifdef _WIN32
typedef lbm_ulong (WINAPI *p_logon)(lbm_handle, char *, char *, lbm_ulong, lbm_ulong, lbm_handle *);
typedef lbm_ulong (WINAPI *p_logoff)(lbm_handle, lbm_handle, lbm_ulong, lbm_ulong);
typedef lbm_ulong (WINAPI *p_send)(lbm_handle, lbm_handle, lbm_message *, lbm_ulong, lbm_ulong);
typedef lbm_ulong (WINAPI *p_senddocs)(lbm_handle, char *, char *, char *, lbm_ulong);
typedef lbm_ulong (WINAPI *p_findnext)(lbm_handle, lbm_handle, char *, char *, lbm_ulong, lbm_ulong, char *);
typedef lbm_ulong (WINAPI *p_read)(lbm_handle, lbm_handle, char *, lbm_ulong, lbm_ulong, lbm_message **);
typedef lbm_ulong (WINAPI *p_save)(lbm_handle, lbm_handle, lbm_message *, lbm_ulong, lbm_ulong, char *);
typedef lbm_ulong (WINAPI *p_delete)(lbm_handle, lbm_handle, char *, lbm_ulong, lbm_ulong);
typedef lbm_ulong (WINAPI *p_free)(void *);
typedef lbm_ulong (WINAPI *p_address)(lbm_handle, lbm_handle, char *, lbm_ulong, char *, lbm_ulong, lbm_recip *,
                                      lbm_ulong, lbm_ulong, lbm_ulong *, lbm_recip **);
typedef lbm_ulong (WINAPI *p_details)(lbm_handle, lbm_handle, lbm_recip *, lbm_ulong, lbm_ulong);
typedef lbm_ulong (WINAPI *p_resolve)(lbm_handle, lbm_handle, char *, lbm_ulong, lbm_ulong, lbm_recip **);
static struct {
    p_logon logon; p_logoff logoff; p_send send; p_senddocs senddocs; p_findnext findnext; p_read read; p_save save;
    p_delete del; p_free free_; p_address address; p_details details; p_resolve resolve;
} w;
#define GET(f, name) do { FARPROC p_ = GetProcAddress(dll, name); memcpy(&w.f, &p_, sizeof p_); } while (0)
#endif

static int which(void)
{
    if (mode != M_UNSET) return mode;
    const char *none = getenv("LBMAPI_NONE");
    if (none && *none && strcmp(none, "0")) return mode = M_NONE;
#ifdef _WIN32
    char buf[8];
    if (GetEnvironmentVariableA("LBMAPI_DIR", buf, sizeof buf)) return mode = M_OWN;
    HMODULE dll = LoadLibraryA("MAPI32.DLL");
    if (!dll) return mode = M_NONE;
    GET(logon, "MAPILogon"); GET(logoff, "MAPILogoff"); GET(send, "MAPISendMail");
    GET(senddocs, "MAPISendDocuments"); GET(findnext, "MAPIFindNext"); GET(read, "MAPIReadMail");
    GET(save, "MAPISaveMail"); GET(del, "MAPIDeleteMail"); GET(free_, "MAPIFreeBuffer"); GET(address, "MAPIAddress");
    GET(details, "MAPIDetails"); GET(resolve, "MAPIResolveName");
    return mode = M_MAPI32;
#else
    return mode = M_OWN;
#endif
}

#ifdef _WIN32
#define MISSING(f) (!w.f ? NO_PROVIDER : 0)
#endif

lbm_ulong lbmc_logon(lbm_handle ui, const char *profile, const char *password, lbm_ulong flags, lbm_handle *session)
{
    *session = 0;
    switch (which()) {
    case M_OWN: return lbm_logon(ui, profile, password, flags, session);
#ifdef _WIN32
    case M_MAPI32: return MISSING(logon) ? NO_PROVIDER : w.logon(ui, (char *)profile, (char *)password, flags, 0, session);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_logoff(lbm_handle session, lbm_handle ui, lbm_ulong flags)
{
    switch (which()) {
    case M_OWN: return lbm_logoff(session, ui, flags);
#ifdef _WIN32
    case M_MAPI32: return MISSING(logoff) ? NO_PROVIDER : w.logoff(session, ui, flags, 0);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_send_mail(lbm_handle session, lbm_handle ui, const lbm_message *msg, lbm_ulong flags)
{
    switch (which()) {
    case M_OWN: return lbm_send_mail(session, ui, msg, flags);
#ifdef _WIN32
    case M_MAPI32: return MISSING(send) ? NO_PROVIDER : w.send(session, ui, (lbm_message *)msg, flags, 0);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_find_next(lbm_handle session, lbm_handle ui, const char *type, const char *seed, lbm_ulong flags,
                         char *id)
{
    id[0] = 0;
    switch (which()) {
    case M_OWN: return lbm_find_next(session, ui, type, seed, flags, id);
#ifdef _WIN32
    case M_MAPI32: return MISSING(findnext) ? NO_PROVIDER : w.findnext(session, ui, (char *)type, (char *)seed, flags, 0, id);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_read_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags, lbm_message **out)
{
    *out = NULL;
    switch (which()) {
    case M_OWN: return lbm_read_mail(session, ui, id, flags, out);
#ifdef _WIN32
    case M_MAPI32: return MISSING(read) ? NO_PROVIDER : w.read(session, ui, (char *)id, flags, 0, out);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_save_mail(lbm_handle session, lbm_handle ui, const lbm_message *msg, lbm_ulong flags, char *id)
{
    switch (which()) {
    case M_OWN: return lbm_save_mail(session, ui, msg, flags, id);
#ifdef _WIN32
    case M_MAPI32: return MISSING(save) ? NO_PROVIDER : w.save(session, ui, (lbm_message *)msg, flags, 0, id);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_delete_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags)
{
    switch (which()) {
    case M_OWN: return lbm_delete_mail(session, ui, id, flags);
#ifdef _WIN32
    case M_MAPI32: return MISSING(del) ? NO_PROVIDER : w.del(session, ui, (char *)id, flags, 0);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_free_buffer(void *p)
{
    if (!p) return LBM_SUCCESS;
    switch (which()) {
    case M_OWN: return lbm_free_buffer(p);
#ifdef _WIN32
    case M_MAPI32: return w.free_ ? w.free_(p) : LBM_SUCCESS;
#endif
    }
    return LBM_SUCCESS;
}

lbm_ulong lbmc_address(lbm_handle session, lbm_handle ui, const char *caption, lbm_ulong nedit, const char *labels,
                       lbm_ulong nrecips, const lbm_recip *recips, lbm_ulong flags, lbm_ulong *nnew, lbm_recip **news)
{
    *nnew = 0;
    *news = NULL;
    switch (which()) {
    case M_OWN: return lbm_address(session, ui, caption, nedit, labels, nrecips, recips, flags, nnew, news);
#ifdef _WIN32
    case M_MAPI32:
        return MISSING(address) ? NO_PROVIDER : w.address(session, ui, (char *)caption, nedit, (char *)labels, nrecips,
                                                          (lbm_recip *)recips, flags, 0, nnew, news);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_details(lbm_handle session, lbm_handle ui, const lbm_recip *recip, lbm_ulong flags)
{
    switch (which()) {
    case M_OWN: return lbm_details(session, ui, recip, flags);
#ifdef _WIN32
    case M_MAPI32: return MISSING(details) ? NO_PROVIDER : w.details(session, ui, (lbm_recip *)recip, flags, 0);
#endif
    }
    return NO_PROVIDER;
}

lbm_ulong lbmc_resolve_name(lbm_handle session, lbm_handle ui, const char *name, lbm_ulong flags, lbm_recip **out)
{
    *out = NULL;
    switch (which()) {
    case M_OWN: return lbm_resolve_name(session, ui, name, flags, out);
#ifdef _WIN32
    case M_MAPI32: return MISSING(resolve) ? NO_PROVIDER : w.resolve(session, ui, (char *)name, flags, 0, out);
#endif
    }
    return NO_PROVIDER;
}
