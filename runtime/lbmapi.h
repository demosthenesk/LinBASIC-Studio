/* lbmapi.h - LinBASIC's own provider of Simple MAPI (phase 7.9): the functions and structures of MAPI.H (Windows SDK),
 * with the messages of a store folder. The MAPI controls of the forms runtime (MSMAPI32.OCX: MAPISession,
 * MAPIMessages) call it on Linux; on Windows they call the MAPI32.DLL of the system, as VB6 does. Built as a DLL
 * (tools/lbmapi: lbmapi.dll, exports MAPILogon ...) it is the mail program of the VB6 tests in the VM, so that the
 * tests of VB6 and of LinBASIC see the same provider.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef LBMAPI_H
#define LBMAPI_H
#include <stdint.h>
#include <stddef.h>
#ifdef __cplusplus
extern "C" {
#endif

typedef uint32_t lbm_ulong;                     /* ULONG, FLAGS (32 bits on every system) */
typedef uintptr_t lbm_handle;                   /* LHANDLE, ULONG_PTR */

typedef struct lbm_recip {                      /* MapiRecipDesc */
    lbm_ulong ulReserved, ulRecipClass;         /* MAPI_ORIG 0, MAPI_TO 1, MAPI_CC 2, MAPI_BCC 3 */
    char *lpszName, *lpszAddress;
    lbm_ulong ulEIDSize;
    void *lpEntryID;
} lbm_recip;
typedef struct lbm_file {                       /* MapiFileDesc */
    lbm_ulong ulReserved, flFlags, nPosition;   /* flFlags: MAPI_OLE 1, MAPI_OLE_STATIC 2 */
    char *lpszPathName, *lpszFileName;
    void *lpFileType;
} lbm_file;
typedef struct lbm_message {                    /* MapiMessage */
    lbm_ulong ulReserved;
    char *lpszSubject, *lpszNoteText, *lpszMessageType, *lpszDateReceived, *lpszConversationID;
    lbm_ulong flFlags;                          /* MAPI_UNREAD 1, MAPI_RECEIPT_REQUESTED 2, MAPI_SENT 4 */
    lbm_recip *lpOriginator;
    lbm_ulong nRecipCount;
    lbm_recip *lpRecips;
    lbm_ulong nFileCount;
    lbm_file *lpFiles;
} lbm_message;

enum {                                          /* flags */
    LBM_LOGON_UI = 0x1, LBM_NEW_SESSION = 0x2, LBM_DIALOG = 0x8, LBM_UNREAD_ONLY = 0x20, LBM_ENVELOPE_ONLY = 0x40,
    LBM_PEEK = 0x80, LBM_GUARANTEE_FIFO = 0x100, LBM_BODY_AS_FILE = 0x200, LBM_AB_NOMODIFY = 0x400,
    LBM_SUPPRESS_ATTACH = 0x800, LBM_FORCE_DOWNLOAD = 0x1000, LBM_LONG_MSGID = 0x4000, LBM_PASSWORD_UI = 0x20000,
    LBM_UNREAD = 0x1, LBM_RECEIPT_REQUESTED = 0x2, LBM_SENT = 0x4
};
enum {                                          /* the results */
    LBM_SUCCESS = 0, LBM_USER_ABORT = 1, LBM_E_FAILURE = 2, LBM_E_LOGIN_FAILURE = 3, LBM_E_DISK_FULL = 4,
    LBM_E_INSUFFICIENT_MEMORY = 5, LBM_E_ACCESS_DENIED = 6, LBM_E_TOO_MANY_SESSIONS = 8, LBM_E_TOO_MANY_FILES = 9,
    LBM_E_TOO_MANY_RECIPIENTS = 10, LBM_E_ATTACHMENT_NOT_FOUND = 11, LBM_E_ATTACHMENT_OPEN_FAILURE = 12,
    LBM_E_ATTACHMENT_WRITE_FAILURE = 13, LBM_E_UNKNOWN_RECIPIENT = 14, LBM_E_BAD_RECIPTYPE = 15,
    LBM_E_NO_MESSAGES = 16, LBM_E_INVALID_MESSAGE = 17, LBM_E_TEXT_TOO_LARGE = 18, LBM_E_INVALID_SESSION = 19,
    LBM_E_TYPE_NOT_SUPPORTED = 20, LBM_E_AMBIGUOUS_RECIPIENT = 21, LBM_E_MESSAGE_IN_USE = 22,
    LBM_E_NETWORK_FAILURE = 23, LBM_E_INVALID_EDITFIELDS = 24, LBM_E_INVALID_RECIPS = 25, LBM_E_NOT_SUPPORTED = 26
};

/* the functions of Simple MAPI (the arguments of MAPI.H; the reserved ones left out) */
lbm_ulong lbm_logon(lbm_handle ui, const char *profile, const char *password, lbm_ulong flags, lbm_handle *session);
lbm_ulong lbm_logoff(lbm_handle session, lbm_handle ui, lbm_ulong flags);
lbm_ulong lbm_send_mail(lbm_handle session, lbm_handle ui, const lbm_message *msg, lbm_ulong flags);
lbm_ulong lbm_send_documents(lbm_handle ui, const char *delims, const char *paths, const char *names);
lbm_ulong lbm_find_next(lbm_handle session, lbm_handle ui, const char *type, const char *seed, lbm_ulong flags,
                        char *id /* 512 bytes */);
lbm_ulong lbm_read_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags, lbm_message **out);
lbm_ulong lbm_save_mail(lbm_handle session, lbm_handle ui, const lbm_message *msg, lbm_ulong flags,
                        char *id /* in: the message to replace or ""; out: its id, 512 bytes */);
lbm_ulong lbm_delete_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags);
lbm_ulong lbm_free_buffer(void *p);
lbm_ulong lbm_address(lbm_handle session, lbm_handle ui, const char *caption, lbm_ulong nedit, const char *labels,
                      lbm_ulong nrecips, const lbm_recip *recips, lbm_ulong flags, lbm_ulong *nnew, lbm_recip **news);
lbm_ulong lbm_details(lbm_handle session, lbm_handle ui, const lbm_recip *recip, lbm_ulong flags);
lbm_ulong lbm_resolve_name(lbm_handle session, lbm_handle ui, const char *name, lbm_ulong flags, lbm_recip **out);

/* where the store is: $LBMAPI_DIR; else (DLL) C:\LBMAPI\store, (Linux) $XDG_DATA_HOME/linbasic-studio/mail */
const char *lbm_store(void);

#ifdef __cplusplus
}
#endif
#endif
