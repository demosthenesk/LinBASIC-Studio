/* ocx_mapi.inc - the MAPI controls (MSMAPI32.OCX: MSMAPI.MAPISession, MSMAPI.MAPIMessages) over Simple MAPI
 * (runtime/mapi.c: LinBASIC's provider lbmapi on Linux, the MAPI32.DLL of the mail program on Windows), as VB6
 * (oracle tests/dev f_mapi1 .. f_mapi4, with lbmapi as the mail program of the VM):
 *   - an error of the provider n is the control's error 32000 + n (no provider: 0x4005 -> 48389); the errors of the
 *     control (32050 ...) have no text of their own ("Application-defined or object-defined error");
 *   - Fetch asks only for the list (MAPIFindNext); a message is read when the program comes to it: MsgIndex, or a
 *     property of it -> its envelope (MAPI_ENVELOPE_ONLY | MAPI_PEEK); MsgNoteText -> its text
 *     (MAPI_SUPPRESS_ATTACH); the attachments, Reply, ReplyAll, Forward, Copy -> all of it. What was read stays
 *     until the next Fetch; AttachmentCount gives the files of the last reading;
 *   - the compose buffer (MsgIndex -1) stays across Fetch; Compose keeps its MsgID (Save after Compose writes
 *     over the message saved before it), Copy keeps it too, Reply / ReplyAll / Forward empty it.
 * Not as VB6, on purpose: after Copy, MsgOrigDisplayName is the name of the originator (VB6: "", though its
 * address is there and Save gives both to the provider).
 * Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <list>
#include "../lbmapi.h"
extern "C" {
lbm_ulong lbmc_logon(lbm_handle ui, const char *profile, const char *password, lbm_ulong flags, lbm_handle *session);
lbm_ulong lbmc_logoff(lbm_handle session, lbm_handle ui, lbm_ulong flags);
lbm_ulong lbmc_send_mail(lbm_handle session, lbm_handle ui, const lbm_message *msg, lbm_ulong flags);
lbm_ulong lbmc_find_next(lbm_handle session, lbm_handle ui, const char *type, const char *seed, lbm_ulong flags,
                         char *id);
lbm_ulong lbmc_read_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags, lbm_message **out);
lbm_ulong lbmc_save_mail(lbm_handle session, lbm_handle ui, const lbm_message *msg, lbm_ulong flags, char *id);
lbm_ulong lbmc_delete_mail(lbm_handle session, lbm_handle ui, const char *id, lbm_ulong flags);
lbm_ulong lbmc_free_buffer(void *p);
lbm_ulong lbmc_address(lbm_handle session, lbm_handle ui, const char *caption, lbm_ulong nedit, const char *labels,
                       lbm_ulong nrecips, const lbm_recip *recips, lbm_ulong flags, lbm_ulong *nnew, lbm_recip **news);
lbm_ulong lbmc_details(lbm_handle session, lbm_handle ui, const lbm_recip *recip, lbm_ulong flags);
lbm_ulong lbmc_resolve_name(lbm_handle session, lbm_handle ui, const char *name, lbm_ulong flags, lbm_recip **out);
}

enum {                                          /* the errors of the control (MAPIErrors) */
    MAP_SESSION_EXIST = 32050, MAP_INVALID_BUFFER = 32051, MAP_INVALID_READ_ACTION = 32052, MAP_NO_SESSION = 32053,
    MAP_INVALID_COMPOSE_ACTION = 32055, MAP_CONTROL_FAILURE = 32056
};
static vbvar mapi_err(int n) { vb_raise(n); return vb_v_empty(); }
/* a result of the provider n: the error 32000 + n; the MAPI32 stub of Windows without a mail program gives the
 * HRESULT E_FAIL (0x80004005), of which the control takes the low word: 48389 (oracle f_mapi1) */
static bool mapi_failed(lbm_ulong r) { if (r) vb_raise((vb_i32)(32000 + (r & 0xFFFF))); return r != 0; }
/* the strings of Simple MAPI: the ANSI code page of the program */
static std::string mapi_ansi(const QString &q)
{
    std::string b;
    for (QChar ch : q) { int c = vb_w2ansi(ch.unicode()); b.push_back((char)(c < 0 ? '?' : c)); }
    return b;
}
static QString mapi_wide(const char *p)
{
    QString s;
    for (; p && *p; p++) s.append(QChar(vb_ansi2w((unsigned char)*p)));
    return s;
}
static vb_i32 mapi_long(vbvar v) { return vb_v2l(v); }

/* ---------------------------------------------------------------- MAPISession */
struct OcxMapiSession : Ocx {
    bool download = true, logonui = true, newsession = false;
    QString user, pass;
    lbm_handle session = 0;
    ~OcxMapiSession() override { if (session) lbmc_logoff(session, 0, 0); }
    QWidget *make(QWidget *) override { return nullptr; }
    void load(const char *const *p) override
    {
        download = propn(p, "DownLoadMail", -1) != 0;
        logonui = propn(p, "LogonUI", -1) != 0;
        newsession = propn(p, "NewSession", 0) != 0;
        prop(p, "UserName", &user);
        prop(p, "Password", &pass);
    }
    vbvar signon()
    {
        if (session) return mapi_err(MAP_SESSION_EXIST);
        std::string u = mapi_ansi(user), pw = mapi_ansi(pass);
        lbm_ulong fl = (logonui ? LBM_LOGON_UI : 0) | (newsession ? LBM_NEW_SESSION : 0) | (download ? LBM_FORCE_DOWNLOAD : 0);
        lbm_handle h = 0;
        if (mapi_failed(lbmc_logon(0, u.c_str(), pw.c_str(), fl, &h))) return vb_v_empty();
        session = h;
        return vb_v_empty();
    }
    vbvar signoff()
    {
        if (!session) return mapi_err(MAP_NO_SESSION);
        lbm_ulong r = lbmc_logoff(session, 0, 0);
        session = 0;
        mapi_failed(r);
        return vb_v_empty();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        auto flag = [&](bool &f) -> vbvar {
            if (!set) return vbool(f);
            vb_bool b = vb_v2bool(val);
            if (!vb_errpend) f = b != 0;
            return vb_v_empty();
        };
        auto str = [&](QString &f) -> vbvar {
            if (!set) return vstr(f);
            vbs s = vb_v2s(val);
            if (!vb_errpend) f = qs(s);
            return vb_v_empty();
        };
        if (n == "action") {
            if (!set) return mapi_err(394);
            vb_i32 a = mapi_long(val);
            if (vb_errpend) return vb_v_empty();
            if (a == 1) return signon();
            if (a == 2) return signoff();
            return mapi_err(380);
        }
        if (n == "sessionid") { if (set) return mapi_err(383); return vb_v_l((vb_i32)session); }
        if (n == "downloadmail") return flag(download);
        if (n == "logonui") return flag(logonui);
        if (n == "newsession") return flag(newsession);
        if (n == "username") return str(user);
        if (n == "password") return str(pass);
        if (n == "signon") return signon();
        if (n == "signoff") return signoff();
        if (n == "aboutbox") return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_mapisession() { return new OcxMapiSession; }

/* ---------------------------------------------------------------- MAPIMessages */
struct MapiRecip { int type = 0; QString name, addr; };
struct MapiFile { QString path, name; vb_i32 pos = 0; int type = 0; };
struct MapiMsg {
    QString id, subject, note, type, date, conv, oname, oaddr;
    bool unread = false, receipt = false, sent = false;
    std::vector<MapiRecip> recips;
    std::vector<MapiFile> files;
    int level = 0;                              /* what was read: 0 nothing, 1 the envelope, 2 + the text, 3 all */
};

/* a MapiMessage of the control's buffer; its strings live as long as it */
struct MapiOut {
    std::list<std::string> keep;
    std::vector<lbm_recip> recips;
    std::vector<lbm_file> files;
    lbm_recip orig{};
    lbm_message m{};
    char *s(const QString &q)                   /* "" -> NULL, as VB6 gives them (oracle f_mapi2) */
    {
        if (q.isEmpty()) return nullptr;
        keep.push_back(mapi_ansi(q));
        return &keep.back()[0];
    }
    lbm_recip recip(const MapiRecip &r)
    {
        lbm_recip x{};
        x.ulRecipClass = (lbm_ulong)r.type;
        x.lpszName = s(r.name);
        x.lpszAddress = s(r.addr);
        return x;
    }
    explicit MapiOut(const MapiMsg &b)
    {
        for (const MapiRecip &r : b.recips) recips.push_back(recip(r));
        for (const MapiFile &f : b.files) {
            lbm_file x{};
            x.flFlags = (lbm_ulong)f.type;
            x.nPosition = (lbm_ulong)f.pos;
            x.lpszPathName = s(f.path);
            x.lpszFileName = s(f.name);
            files.push_back(x);
        }
        m.lpszSubject = s(b.subject);
        m.lpszNoteText = s(b.note);
        m.lpszMessageType = s(b.type);
        m.lpszDateReceived = s(b.date);
        m.lpszConversationID = s(b.conv);
        m.flFlags = (b.receipt ? LBM_RECEIPT_REQUESTED : 0) | (b.sent ? LBM_SENT : 0);
        if (!b.oname.isEmpty() || !b.oaddr.isEmpty()) {
            orig.ulRecipClass = 0;
            orig.lpszName = s(b.oname);
            orig.lpszAddress = s(b.oaddr);
            m.lpOriginator = &orig;
        }
        m.nRecipCount = (lbm_ulong)recips.size();
        m.lpRecips = recips.empty() ? nullptr : recips.data();
        m.nFileCount = (lbm_ulong)files.size();
        m.lpFiles = files.empty() ? nullptr : files.data();
    }
};

struct OcxMapiMessages : Ocx {
    QString caption, label, fetchtype;
    vb_i16 editfields = 1;
    bool modifiable = false, resolveui = false, sorted = false, unreadonly = false;
    vb_i32 sessionid = 0, msgindex = 0, recipindex = 0, attindex = 0;
    std::vector<MapiMsg> list;                  /* what Fetch gave */
    MapiMsg buf;                                /* the compose buffer (MsgIndex -1) */
    MapiMsg none;                               /* a MsgIndex without a message */

    QWidget *make(QWidget *) override { return nullptr; }
    void load(const char *const *p) override
    {
        prop(p, "AddressCaption", &caption);
        prop(p, "AddressLabel", &label);
        prop(p, "FetchMsgType", &fetchtype);
        editfields = (vb_i16)propn(p, "AddressEditFieldCount", 1);
        modifiable = propn(p, "AddressModifiable", 0) != 0;
        resolveui = propn(p, "AddressResolveUI", 0) != 0;
        sorted = propn(p, "FetchSorted", 0) != 0;
        unreadonly = propn(p, "FetchUnreadOnly", 0) != 0;
    }
    lbm_handle sess() const { return (lbm_handle)(vb_u32)sessionid; }
    bool composing() const { return msgindex == -1; }
    bool valid() const { return msgindex >= 0 && msgindex < (vb_i32)list.size(); }
    MapiMsg &cur() { return composing() ? buf : valid() ? list[(size_t)msgindex] : (none = MapiMsg(), none); }

    /* the message at MsgIndex read up to a level (1 envelope, 2 text, 3 all); false: an error is raised */
    bool read_one(MapiMsg &m, int want)
    {
        lbm_ulong fl = want == 1 ? (LBM_ENVELOPE_ONLY | LBM_PEEK) : want == 2 ? LBM_SUPPRESS_ATTACH : 0;
        std::string id = mapi_ansi(m.id);
        lbm_message *out = nullptr;
        if (mapi_failed(lbmc_read_mail(sess(), 0, id.c_str(), fl, &out))) return false;
        if (!out) return true;
        m.subject = mapi_wide(out->lpszSubject);
        if (want >= 2) m.note = mapi_wide(out->lpszNoteText);
        m.type = mapi_wide(out->lpszMessageType);
        m.date = mapi_wide(out->lpszDateReceived);
        m.conv = mapi_wide(out->lpszConversationID);
        m.unread = (out->flFlags & LBM_UNREAD) && want == 1;   /* read without MAPI_PEEK: read now */
        m.receipt = (out->flFlags & LBM_RECEIPT_REQUESTED) != 0;
        m.sent = (out->flFlags & LBM_SENT) != 0;
        m.oname = out->lpOriginator ? mapi_wide(out->lpOriginator->lpszName) : QString();
        m.oaddr = out->lpOriginator ? mapi_wide(out->lpOriginator->lpszAddress) : QString();
        m.recips.clear();
        for (lbm_ulong i = 0; i < out->nRecipCount && out->lpRecips; i++)
            m.recips.push_back({ (int)out->lpRecips[i].ulRecipClass, mapi_wide(out->lpRecips[i].lpszName),
                                 mapi_wide(out->lpRecips[i].lpszAddress) });
        m.files.clear();
        for (lbm_ulong i = 0; i < out->nFileCount && out->lpFiles; i++)
            m.files.push_back({ mapi_wide(out->lpFiles[i].lpszPathName), mapi_wide(out->lpFiles[i].lpszFileName),
                                (vb_i32)out->lpFiles[i].nPosition, (int)out->lpFiles[i].flFlags });
        m.level = want;
        lbmc_free_buffer(out);
        return true;
    }
    bool need(int want)
    {
        if (!valid()) return true;
        MapiMsg &m = list[(size_t)msgindex];
        if (m.level == 0 && !read_one(m, 1)) return false;
        if (m.level < want && !read_one(m, want)) return false;
        return true;
    }

    /* ------------------------------------------------ the actions */
    vbvar fetch()
    {
        if (!sessionid) return mapi_err(MAP_NO_SESSION);
        list.clear();
        msgindex = recipindex = attindex = 0;
        std::string t = mapi_ansi(fetchtype);
        lbm_ulong fl = (unreadonly ? LBM_UNREAD_ONLY : 0) | (sorted ? LBM_GUARANTEE_FIFO : 0);
        char seed[512] = "", id[512];
        for (int k = 0; k < 1000000; k++) {
            lbm_ulong r = lbmc_find_next(sess(), 0, t.c_str(), seed, fl, id);
            if (r == LBM_E_NO_MESSAGES) break;
            if (r) { list.clear(); mapi_failed(r); return vb_v_empty(); }
            if (!id[0] || !strcmp(id, seed)) break;   /* a provider that gives the seed again: the end */
            MapiMsg m;
            m.id = mapi_wide(id);
            list.push_back(m);
            snprintf(seed, sizeof seed, "%s", id);
        }
        return vb_v_empty();
    }
    vbvar compose()
    {
        if (!sessionid) return mapi_err(MAP_NO_SESSION);
        MapiMsg b;
        b.id = buf.id;                          /* Compose keeps the MsgID (oracle f_mapi3) */
        buf = b;
        msgindex = -1;
        recipindex = attindex = 0;
        return vb_v_empty();
    }
    vbvar send(bool dialog)
    {
        if (!composing()) return mapi_err(MAP_INVALID_READ_ACTION);
        if (!sessionid) return mapi_err(MAP_NO_SESSION);
        MapiOut o(buf);
        mapi_failed(lbmc_send_mail(sess(), 0, &o.m, dialog ? LBM_DIALOG : 0));
        return vb_v_empty();
    }
    vbvar save()
    {
        if (!composing()) return mapi_err(MAP_INVALID_READ_ACTION);
        if (!sessionid) return mapi_err(MAP_NO_SESSION);
        MapiOut o(buf);
        char id[512];
        snprintf(id, sizeof id, "%s", mapi_ansi(buf.id).c_str());
        if (mapi_failed(lbmc_save_mail(sess(), 0, &o.m, 0, id))) return vb_v_empty();
        buf.id = mapi_wide(id);
        return vb_v_empty();
    }
    /* Copy 0, Reply 1, ReplyAll 2, Forward 3: the message at MsgIndex into the compose buffer */
    vbvar copy(int how)
    {
        if (composing()) return mapi_err(MAP_INVALID_COMPOSE_ACTION);
        if (!valid()) return mapi_err(MAP_CONTROL_FAILURE);
        if (!need(3)) return vb_v_empty();
        MapiMsg b = list[(size_t)msgindex];
        b.level = 0;
        b.unread = false;
        if (how) {
            MapiRecip from{ 1, b.oname, b.oaddr };
            std::vector<MapiRecip> rs;
            if (how != 3) rs.push_back(from);
            if (how == 2) for (MapiRecip r : b.recips) { r.type = 1; rs.push_back(r); }
            b.recips = rs;
            b.subject = (how == 3 ? "FW: " : "RE: ") + b.subject;
            b.oname.clear(); b.oaddr.clear();
            b.id.clear();
        }
        buf = b;
        msgindex = -1;
        recipindex = attindex = 0;
        return vb_v_empty();
    }
    vbvar del(vb_i32 argc, vbvar *argv)
    {
        vb_i32 what = given(argc, argv, 0) ? mapi_long(argv[0]) : 0;
        if (vb_errpend) return vb_v_empty();
        if (what < 0 || what > 2) return mapi_err(380);
        if (what == 0) {
            if (composing()) return mapi_err(MAP_INVALID_COMPOSE_ACTION);
            if (!sessionid) return mapi_err(MAP_NO_SESSION);
            if (!valid()) return mapi_err(MAP_CONTROL_FAILURE);
            std::string id = mapi_ansi(list[(size_t)msgindex].id);
            if (mapi_failed(lbmc_delete_mail(sess(), 0, id.c_str(), 0))) return vb_v_empty();
            list.erase(list.begin() + msgindex);
            if (msgindex >= (vb_i32)list.size()) msgindex = std::max(0, (int)list.size() - 1);
            recipindex = attindex = 0;
            return vb_v_empty();
        }
        if (!composing()) return mapi_err(MAP_INVALID_READ_ACTION);
        if (what == 1) {
            if (recipindex < (vb_i32)buf.recips.size()) buf.recips.erase(buf.recips.begin() + recipindex);
            if (recipindex >= (vb_i32)buf.recips.size()) recipindex = std::max(0, (int)buf.recips.size() - 1);
        } else {
            if (attindex < (vb_i32)buf.files.size()) buf.files.erase(buf.files.begin() + attindex);
            if (attindex >= (vb_i32)buf.files.size()) attindex = std::max(0, (int)buf.files.size() - 1);
        }
        return vb_v_empty();
    }
    vbvar resolve()
    {
        if (!composing()) return vb_v_empty();  /* a message read: nothing (oracle f_mapi1, f_mapi4) */
        if (!sessionid) return mapi_err(MAP_NO_SESSION);
        if (recipindex >= (vb_i32)buf.recips.size()) return vb_v_empty();
        MapiRecip &r = buf.recips[(size_t)recipindex];
        std::string nm = mapi_ansi(r.name);
        lbm_recip *out = nullptr;
        lbm_ulong fl = (modifiable ? 0 : LBM_AB_NOMODIFY) | (resolveui ? LBM_DIALOG : 0);
        if (mapi_failed(lbmc_resolve_name(sess(), 0, nm.c_str(), fl, &out))) return vb_v_empty();
        if (out) {
            r.name = mapi_wide(out->lpszName);
            r.addr = mapi_wide(out->lpszAddress);
            lbmc_free_buffer(out);
        }
        return vb_v_empty();
    }
    vbvar show(bool details)
    {
        if (!composing()) return mapi_err(MAP_INVALID_READ_ACTION);
        if (!sessionid) return mapi_err(MAP_NO_SESSION);
        lbm_ulong fl = modifiable ? 0 : LBM_AB_NOMODIFY;
        MapiOut o(buf);
        if (details) {
            if (recipindex >= (vb_i32)buf.recips.size()) buf.recips.resize((size_t)recipindex + 1);
            lbm_recip rc = o.recip(buf.recips[(size_t)recipindex]);
            mapi_failed(lbmc_details(sess(), 0, &rc, fl));
            return vb_v_empty();
        }
        lbm_ulong nn = 0;
        lbm_recip *news = nullptr;
        if (mapi_failed(lbmc_address(sess(), 0, o.s(caption), (lbm_ulong)editfields, o.s(label), o.m.nRecipCount,
                                     o.m.lpRecips, fl, &nn, &news))) return vb_v_empty();
        buf.recips.clear();
        for (lbm_ulong i = 0; i < nn && news; i++)
            buf.recips.push_back({ (int)news[i].ulRecipClass, mapi_wide(news[i].lpszName), mapi_wide(news[i].lpszAddress) });
        lbmc_free_buffer(news);
        return vb_v_empty();
    }
    vbvar action(vb_i32 a)
    {
        vbvar no[1] = { vb_v_l(0) }, one[1] = { vb_v_l(1) }, two[1] = { vb_v_l(2) };
        switch (a) {
        case 1: return fetch();
        case 2: return send(true);
        case 3: return send(false);
        case 4: return save();
        case 5: return copy(0);
        case 6: return compose();
        case 7: return copy(1);
        case 8: return copy(2);
        case 9: return copy(3);
        case 10: return del(1, no);
        case 11: return show(false);
        case 12: return show(true);
        case 13: return resolve();
        case 14: return del(1, one);
        case 15: return del(1, two);
        }
        return mapi_err(380);
    }

    /* ------------------------------------------------ the recipients and attachments of the message */
    MapiRecip *recip(bool grow)
    {
        MapiMsg &m = cur();
        if (recipindex >= (vb_i32)m.recips.size()) {
            if (!grow || !composing()) return nullptr;
            m.recips.resize((size_t)recipindex + 1);   /* the compose buffer grows to the index (oracle f_mapi1) */
        }
        return &m.recips[(size_t)recipindex];
    }
    MapiFile *file(bool grow)
    {
        MapiMsg &m = cur();
        if (attindex >= (vb_i32)m.files.size()) {
            if (!grow || !composing()) return nullptr;
            m.files.resize((size_t)attindex + 1);
        }
        return &m.files[(size_t)attindex];
    }

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        auto flag = [&](bool &f) -> vbvar {
            if (!set) return vbool(f);
            vb_bool b = vb_v2bool(val);
            if (!vb_errpend) f = b != 0;
            return vb_v_empty();
        };
        auto str = [&](QString &f) -> vbvar {
            if (!set) return vstr(f);
            vbs s = vb_v2s(val);
            if (!vb_errpend) f = qs(s);
            return vb_v_empty();
        };
        /* the methods */
        if (n == "fetch") return fetch();
        if (n == "compose") return compose();
        if (n == "send") {
            bool d = given(argc, argv, 0) && vb_v2bool(argv[0]) != 0;
            return vb_errpend ? vb_v_empty() : send(d);
        }
        if (n == "save") return save();
        if (n == "copy") return copy(0);
        if (n == "reply") return copy(1);
        if (n == "replyall") return copy(2);
        if (n == "forward") return copy(3);
        if (n == "delete") return del(argc, argv);
        if (n == "resolvename") return resolve();
        if (n == "show") {
            bool d = given(argc, argv, 0) && vb_v2bool(argv[0]) != 0;
            return vb_errpend ? vb_v_empty() : show(d);
        }
        if (n == "aboutbox") return vb_v_empty();
        if (n == "action") {
            if (!set) return mapi_err(394);
            vb_i32 a = mapi_long(val);
            return vb_errpend ? vb_v_empty() : action(a);
        }
        /* the properties of the control */
        if (n == "addresscaption") return str(caption);
        if (n == "addresslabel") return str(label);
        if (n == "fetchmsgtype") return str(fetchtype);
        if (n == "addressmodifiable") return flag(modifiable);
        if (n == "addressresolveui") return flag(resolveui);
        if (n == "fetchsorted") return flag(sorted);
        if (n == "fetchunreadonly") return flag(unreadonly);
        if (n == "addresseditfieldcount") {
            if (!set) return vb_v_i(editfields);
            vb_i32 v = mapi_long(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0 || v > 4) return mapi_err(32000 + LBM_E_INVALID_EDITFIELDS);
            editfields = (vb_i16)v;
            return vb_v_empty();
        }
        if (n == "sessionid") {
            if (!set) return vb_v_l(sessionid);
            vb_i32 v = mapi_long(val);
            if (!vb_errpend) sessionid = v;
            return vb_v_empty();
        }
        if (n == "msgcount") { if (set) return mapi_err(383); return vb_v_l((vb_i32)list.size()); }
        if (n == "msgindex") {
            if (!set) return vb_v_l(msgindex);
            vb_i32 v = mapi_long(val);
            if (vb_errpend) return vb_v_empty();
            if (v < -1 || v >= (vb_i32)list.size()) return mapi_err(380);
            msgindex = v;
            recipindex = attindex = 0;
            need(1);                            /* its envelope (oracle f_mapi4) */
            return vb_v_empty();
        }
        /* the properties of the message at MsgIndex: read only ... */
        static const char *const ro[] = { "msgid", "msgorigaddress", "msgorigdisplayname", "msgread", "msgdatereceived",
                                          "msgsent", "recipcount", "attachmentcount" };
        for (const char *r : ro) if (set && n == r) return mapi_err(383);
        static const char *const msgprops[] = { "msgid", "msgorigaddress", "msgorigdisplayname", "msgread",
            "msgdatereceived", "msgsent", "recipcount", "attachmentcount", "msgsubject", "msgnotetext", "msgtype",
            "msgconversationid", "msgreceiptrequested", "recipindex", "recipaddress", "recipdisplayname", "reciptype",
            "attachmentindex", "attachmentname", "attachmentpathname", "attachmentposition", "attachmenttype" };
        bool known = false;
        for (const char *r : msgprops) if (n == r) known = true;
        if (!known) { *found = 0; return vb_v_empty(); }
        bool att = n.compare(0, 10, "attachment") == 0 && n != "attachmentcount" && n != "attachmentindex";
        if (!need(n == "msgnotetext" ? 2 : att && !set ? 3 : 1)) return vb_v_empty();
        bool rw = composing();
        /* ... and those that a message read does not let change (32051) */
        if (set && !rw && n != "recipindex" && n != "attachmentindex") return mapi_err(MAP_INVALID_BUFFER);
        MapiMsg &m = cur();
        if (n == "msgid") return vstr(m.id);
        if (n == "msgorigaddress") return vstr(m.oaddr);
        if (n == "msgorigdisplayname") return vstr(m.oname);
        if (n == "msgread") return vbool(rw || (valid() && !m.unread));
        if (n == "msgdatereceived") return vstr(m.date);
        if (n == "msgsent") return vbool(m.sent);
        if (n == "recipcount") return vb_v_l((vb_i32)m.recips.size());
        if (n == "attachmentcount") return vb_v_l((vb_i32)m.files.size());
        if (n == "msgsubject") return str(m.subject);
        if (n == "msgnotetext") return str(m.note);
        if (n == "msgtype") return str(m.type);
        if (n == "msgconversationid") return str(m.conv);
        if (n == "msgreceiptrequested") return flag(m.receipt);
        if (n == "recipindex" || n == "attachmentindex") {
            bool r = n == "recipindex";
            vb_i32 &ix = r ? recipindex : attindex;
            if (!set) return vb_v_l(ix);
            vb_i32 v = mapi_long(val);
            if (vb_errpend) return vb_v_empty();
            size_t cnt = r ? m.recips.size() : m.files.size();
            if (v < 0 || (!rw && v >= (vb_i32)cnt)) return mapi_err(380);
            ix = v;
            if (rw) { if (r) recip(true); else file(true); }   /* the compose buffer grows to it */
            return vb_v_empty();
        }
        if (n.compare(0, 5, "recip") == 0) {
            MapiRecip *r = recip(true);
            MapiRecip empty;
            if (!r) r = &empty;
            if (n == "recipaddress") return str(r->addr);
            if (n == "recipdisplayname") {
                vbvar x = str(r->name);
                if (set && !vb_errpend && r->type == 0) r->type = 1;   /* a name makes it a To (oracle f_mapi1) */
                return x;
            }
            if (!set) return vb_v_i((vb_i16)r->type);
            vb_i32 v = mapi_long(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0 || v > 3) return mapi_err(380);
            r->type = (int)v;
            return vb_v_empty();
        }
        MapiFile *f = file(true);
        MapiFile empty;
        if (!f) f = &empty;
        if (n == "attachmentname") return str(f->name);
        if (n == "attachmentpathname") return str(f->path);
        if (n == "attachmentposition") {
            if (!set) return vb_v_l(f->pos);
            vb_i32 v = mapi_long(val);
            if (!vb_errpend) f->pos = v;
            return vb_v_empty();
        }
        if (!set) return vb_v_i((vb_i16)f->type);
        vb_i32 v = mapi_long(val);
        if (vb_errpend) return vb_v_empty();
        if (v < 0 || v > 2) return mapi_err(380);
        f->type = (int)v;
        return vb_v_empty();
    }
};
static Ocx *new_mapimessages() { return new OcxMapiMessages; }
