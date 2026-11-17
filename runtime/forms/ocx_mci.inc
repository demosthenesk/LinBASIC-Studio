/* ocx_mci.inc - the Multimedia control (MCI32.OCX, MMControl): the nine buttons and the MCI
 * commands they send (Command = "Open", "Play" ...), through vb_mci_send (mci.c: Linux has its
 * own MCI, Windows the real winmm). Also the pictures of video devices: in the window of
 * hWndDisplay or in a window of their own. Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

static QWidget *widget_of_hwnd(vb_i32 h);

/* ---------------------------------------------------------------- the pictures of video devices */
struct MciView {
    QPointer<QLabel> label;                     /* the picture, in the target window */
    QPointer<QWidget> own;                      /* the default window of the device */
    vb_i64 hwnd = -1;
};
static std::map<vb_i32, MciView> mci_views;
static QTimer *mci_timer;

static void mci_tick()
{
    vb_mci_poll();
    std::set<vb_i32> alive;
    for (vb_i32 id = vb_mci_next_video(0); id; id = vb_mci_next_video(id)) {
        alive.insert(id);
        vb_mci_frameinfo f; memset(&f, 0, sizeof f);
        if (!vb_mci_frame(id, &f)) continue;
#ifdef _WIN32
        (void)f; continue;                      /* Windows: the real MCI draws */
#endif
        MciView &v = mci_views[id];
        if (v.hwnd != f.hwnd) {
            if (v.label) v.label->deleteLater();
            v.label = nullptr;
            v.hwnd = f.hwnd;
        }
        if (!f.data && !f.show) continue;
        if (!v.label) {
            QWidget *host = f.hwnd ? widget_of_hwnd((vb_i32)f.hwnd) : nullptr;
            if (!host) {
                if (!f.show) continue;
                if (!v.own) {
                    v.own = new QWidget(nullptr, Qt::Window);
                    v.own->setWindowTitle(QString::fromUtf8(f.title));
                    v.own->resize(std::max(f.w, 16), std::max(f.h, 16));
                    v.own->setAttribute(Qt::WA_DeleteOnClose, false);
                }
                host = v.own;
                v.own->show();
            }
            v.label = new QLabel(host);
            v.label->setAttribute(Qt::WA_TransparentForMouseEvents);
            v.label->move(0, 0);
            v.label->resize(f.w, f.h);
            v.label->show();
        }
        if (f.data && f.fresh) {
            QImage img(f.data, f.w, f.h, f.stride, QImage::Format_RGB32);
            v.label->setPixmap(QPixmap::fromImage(img.copy()));
            v.label->resize(f.w, f.h);
        }
    }
    for (auto it = mci_views.begin(); it != mci_views.end();) {
        if (!alive.count(it->first)) {
            if (it->second.label) it->second.label->deleteLater();
            if (it->second.own) it->second.own->deleteLater();
            it = mci_views.erase(it);
        } else ++it;
    }
    if (!vb_mci_active() && mci_views.empty() && mci_timer) mci_timer->stop();
}
static void mci_wake()
{
    if (!mci_timer) {
        mci_timer = new QTimer();
        QObject::connect(mci_timer, &QTimer::timeout, []() { mci_tick(); });
    }
    if (!mci_timer->isActive()) mci_timer->start(15);
}

/* ---------------------------------------------------------------- the control */
enum { MB_PREV, MB_NEXT, MB_PLAY, MB_PAUSE, MB_BACK, MB_STEP, MB_STOP, MB_RECORD, MB_EJECT, MB_N };
static const char *const mb_name[MB_N] = { "prev", "next", "play", "pause", "back", "step", "stop", "record", "eject" };
static const char *const mb_cmd[MB_N] = { "Prev", "Next", "Play", "Pause", "Back", "Step", "Stop", "Record", "Eject" };

/* the picture of a button: the glyphs of the Multimedia control */
static QIcon mb_icon(int b)
{
    QPixmap pm(16, 16);
    pm.fill(Qt::transparent);
    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing, false);
    p.setPen(Qt::NoPen);
    p.setBrush(b == MB_RECORD ? QColor(160, 0, 0) : QColor(0, 0, 0));
    auto tri = [&](int x, int y, int w, int h, bool right) {
        QPolygon pg;
        if (right) pg << QPoint(x, y) << QPoint(x + w, y + h / 2) << QPoint(x, y + h);
        else pg << QPoint(x + w, y) << QPoint(x, y + h / 2) << QPoint(x + w, y + h);
        p.drawPolygon(pg);
    };
    switch (b) {
    case MB_PREV: p.drawRect(2, 4, 2, 9); tri(4, 4, 5, 9, false); tri(9, 4, 5, 9, false); break;
    case MB_NEXT: tri(2, 4, 5, 9, true); tri(7, 4, 5, 9, true); p.drawRect(12, 4, 2, 9); break;
    case MB_PLAY: tri(5, 3, 7, 11, true); break;
    case MB_PAUSE: p.drawRect(4, 4, 3, 9); p.drawRect(9, 4, 3, 9); break;
    case MB_BACK: p.drawRect(3, 4, 2, 9); tri(6, 4, 6, 9, false); break;
    case MB_STEP: tri(4, 4, 6, 9, true); p.drawRect(11, 4, 2, 9); break;
    case MB_STOP: p.drawRect(4, 4, 8, 8); break;
    case MB_RECORD: p.drawEllipse(4, 4, 8, 8); break;
    case MB_EJECT: { QPolygon pg; pg << QPoint(8, 3) << QPoint(13, 9) << QPoint(3, 9); p.drawPolygon(pg); p.drawRect(3, 11, 11, 2); break; }
    }
    return QIcon(pm);
}

struct OcxMMControl;
class MciWidget : public QWidget {
public:
    OcxMMControl *mm = nullptr;
    QPushButton *btn[MB_N] = {};
    explicit MciWidget(QWidget *parent) : QWidget(parent) {}
    void relayout();
protected:
    void resizeEvent(QResizeEvent *) override { relayout(); }
};

struct OcxMMControl : Ocx {
    QPointer<MciWidget> w;
    QString devtype, filename, alias, command, errmsg;
    bool autoenable = true, notify = false, wait = false, shareable = false, silent = false, enabled = true;
    bool vis[MB_N], en[MB_N];
    int orientation = 0, recmode = 1, border = 1, tf = 0, updint = 1000, error = 257, devid = 0, notifyvalue = 0;
    vb_i32 from = 0, to = 0, frames = 1, track = 1;
    bool from_set = false, to_set = false;
    bool notify_armed = false, wait_armed = false, play_pending = false;
    int stop_rewrite = -1;
    vb_i64 hwnddisp = 0;
    QTimer *upd = nullptr;
    int lastmode = 524;
    static int next_alias;

    ~OcxMMControl() override
    {
        vb_mci_forget(this);
        if (is_open()) send("close " + alias);
    }
    OcxMMControl()
    {
        for (int i = 0; i < MB_N; i++) { vis[i] = true; en[i] = false; }
        set_error(257);                         /* VB6: a new control has no device: 257 (invalid device ID) */
    }
    void set_error(int e)
    {
        error = e;
        char t[1024] = "";
        if (e) vb_mci_error_text(e, t, sizeof t);
        errmsg = QString::fromUtf8(t);
    }
    QWidget *make(QWidget *parent) override
    {
        w = new MciWidget(parent);
        w->mm = this;
        for (int i = 0; i < MB_N; i++) {
            auto *b = new QPushButton(w);
            b->setIcon(mb_icon(i));
            b->setFocusPolicy(Qt::TabFocus);
            QObject::connect(b, &QPushButton::clicked, [this, i]() { clicked(i); });
            w->btn[i] = b;
        }
        upd = new QTimer(w);
        QObject::connect(upd, &QTimer::timeout, [this]() { status_tick(); });
        sync();
        return w;
    }
    void load(const char *const *p) override
    {
        for (int i = 0; i < MB_N; i++) {
            std::string b = mb_cmd[i];
            vis[i] = propn(p, (b + "Visible").c_str(), 1) != 0;
            en[i] = propn(p, (b + "Enabled").c_str(), 0) != 0;
        }
        autoenable = propn(p, "AutoEnable", 1) != 0;
        orientation = (int)propn(p, "Orientation", 0);
        border = (int)propn(p, "BorderStyle", 1);
        recmode = (int)propn(p, "RecordMode", 1);
        shareable = propn(p, "Shareable", 0) != 0;
        silent = propn(p, "Silent", 0) != 0;
        updint = (int)propn(p, "UpdateInterval", 1000);
        frames = (int)propn(p, "Frames", 1);
        enabled = propn(p, "Enabled", 1) != 0;
        prop(p, "DeviceType", &devtype);
        prop(p, "FileName", &filename);
        sync();
    }
    bool is_open() const { return !alias.isEmpty(); }

    /* a command of MCI for this device; the error goes to Error / ErrorMessage */
    int send(const QString &cmd, QString *ret = nullptr, bool use_notify = false)
    {
        QByteArray c = cmd.toUtf8();
        char r[1024] = "";
        vb_i32 e = vb_mci_send(c.constData(), r, sizeof r, use_notify ? &OcxMMControl::notified : nullptr, use_notify ? this : nullptr);
        if (use_notify) mci_wake();
        if (ret) *ret = QString::fromUtf8(r);
        return e;
    }
    QString q(const QString &what)              /* status ... */
    {
        QString r;
        if (!is_open() || send("status " + alias + " " + what, &r)) return QString();
        return r;
    }
    static void notified(void *ctx, vb_i32, vb_i32 code)
    {
        OcxMMControl *m = (OcxMMControl *)ctx;
        if (code == 4 && m->stop_rewrite >= 0) {   /* the Play a Stop ended */
            int r = m->stop_rewrite;
            m->stop_rewrite = -1;
            m->play_pending = false;
            if (!r) return;
            code = 1;
        }
        if (code != 2) m->play_pending = false;
        m->notifyvalue = code;
        vbvar a = vb_v_i((vb_i16)code);
        fire_ctl(m->c, "done", 1, &a);
        m->sync();
    }
    int mode()
    {
        if (!is_open()) return 524;
        QString s = q("mode");
        static const char *const names[] = { "not ready", "stopped", "playing", "recording", "seeking", "paused", "open" };
        for (int i = 0; i < 7; i++) if (s == names[i]) return 524 + i;
        return 524;
    }
    static const char *tf_word(int t)
    {
        switch (t) {
        case 0: return "ms"; case 1: return "hms"; case 2: return "msf"; case 3: return "frames"; case 4: return "smpte 24";
        case 5: return "smpte 25"; case 6: return "smpte 30"; case 7: return "smpte 30 drop"; case 8: return "bytes";
        case 9: return "samples"; case 10: return "tmsf";
        }
        return "ms";
    }
    /* a time of MCI as the Long of the control: the packed formats (MCI_MAKE_TMSF ...) */
    vb_i32 time_value(const QString &s)
    {
        if (!s.contains(':')) return (vb_i32)s.toLongLong();
        QStringList f = s.split(':');
        vb_u32 v = 0;
        for (int i = 0; i < f.size() && i < 4; i++) v |= (vb_u32)(f[i].toUInt() & 0xFF) << (8 * i);
        return (vb_i32)v;
    }
    QString time_text(vb_i32 v)
    {
        switch (tf) {
        case 1: return QString("%1:%2:%3").arg(v & 0xFF).arg((v >> 8) & 0xFF).arg((v >> 16) & 0xFF);
        case 2: return QString("%1:%2:%3").arg(v & 0xFF).arg((v >> 8) & 0xFF).arg((v >> 16) & 0xFF);
        case 4: case 5: case 6: case 7: case 10:
            return QString("%1:%2:%3:%4").arg(v & 0xFF).arg((v >> 8) & 0xFF).arg((v >> 16) & 0xFF).arg((vb_u32)v >> 24);
        }
        return QString::number(v);
    }
    bool caps_valid = false;
    std::map<std::string, bool> caps;            /* VB6 keeps them after Close (oracle f_mci) */
    bool cap(const char *what)
    {
        if (!is_open()) { auto it = caps.find(what); return it != caps.end() && it->second; }
        QString r;
        bool v = !send(QString("capability ") + alias + " " + what, &r) && r == "true";
        caps[what] = v;
        return v;
    }

    /* the buttons that work now (AutoEnable) */
    void sync()
    {
        if (!w) return;
        bool live[MB_N];
        if (autoenable) {
            int m = mode();
            bool open = m != 524;
            bool play = open && (m == 526 || m == 527 || m == 528);
            bool canplay = open && cap("can play"), canrec = open && cap("can record"), caneject = open && cap("can eject");
            bool canstep = open && devtype_has_step();
            live[MB_PREV] = open && !play ? true : open;
            live[MB_NEXT] = open;
            live[MB_PLAY] = canplay && !play;
            live[MB_PAUSE] = play;
            live[MB_BACK] = canstep && !play;
            live[MB_STEP] = canstep && !play;
            live[MB_STOP] = play;
            live[MB_RECORD] = canrec && !play;
            live[MB_EJECT] = caneject;
            lastmode = m;
        } else for (int i = 0; i < MB_N; i++) live[i] = en[i];
        for (int i = 0; i < MB_N; i++) {
            w->btn[i]->setVisible(vis[i]);
            w->btn[i]->setEnabled(enabled && live[i]);
        }
        w->relayout();
    }
    bool devtype_has_step() { return cap("has video"); }

    void clicked(int b)
    {
        std::string n = mb_name[b];
        vbvar cancel = vb_v_i(0);
        fire_ctl(c, (n + "click").c_str(), 1, &cancel);
        bool stop = vb_v2l(cancel) != 0;
        vb_vfree(&cancel);
        if (stop) return;
        vb_i32 e = run(mb_cmd[b], false);
        vbvar a = vb_v_l(e);
        fire_ctl(c, (n + "completed").c_str(), 1, &a);
    }
    void status_tick()
    {
        if (!is_open() || updint <= 0) { upd->stop(); return; }
        fire_ctl(c, "statusupdate");
        if (autoenable && mode() != lastmode) sync();
    }
    void restart_updates()
    {
        if (!upd) return;
        if (is_open() && updint > 0) upd->start(updint); else upd->stop();
    }

    /* Command: returns the MCI error (0: done) */
    vb_i32 run(const QString &cmdname, bool raise)
    {
        QString cmd = cmdname.trimmed().toLower();
        QString flags;
        /* VB6 (oracle f_mci): Wait is for the next command only; Notify stays, every command uses it,
         * except Prev / Next, which notify only when Notify was set again since the last command */
        bool prevnext = cmd == "prev" || cmd == "next";
        bool use_notify = prevnext ? notify_armed : notify;
        bool armed = notify_armed;
        if (wait_armed) flags += " wait";
        if (use_notify) flags += " notify";
        wait_armed = false; notify_armed = false;
        /* a Stop while a Play waits to notify: VB6 gives that Play 1 if Notify was set again before
         * the Stop, otherwise no Done at all (oracle f_mci A5, A6) */
        if (cmd == "stop" && play_pending) stop_rewrite = armed ? 1 : 0;
        vb_i32 e = 0;
        QString r;
        auto fromto = [&]() {
            QString s;
            if (from_set) s += " from " + time_text(from);
            if (to_set) s += " to " + time_text(to);
            return s;
        };
        if (cmd == "open") {
            QString type = devtype.trimmed(), file = filename.trimmed();
            if (type.isEmpty() && file.isEmpty()) { set_error(257); sync(); return 257; }
            QString old = alias;                /* the device open before: closed after the new one opens */
            /* a waveaudio file that is not there: a new recording (VB6 oracle f_mci) */
            if (!file.isEmpty() && !type.compare("waveaudio", Qt::CaseInsensitive) && !QFileInfo::exists(native(file)))
                file.clear();
            QString al = QString("mmc%1").arg(++next_alias);
            QString s = "open ";
            if (!file.isEmpty()) s += "\"" + file + "\"" + (type.isEmpty() ? QString() : " type " + type);
            else s += type.isEmpty() ? QString("\"\"") : type;
            if (file.isEmpty() && !type.compare("waveaudio", Qt::CaseInsensitive)) s = "open new type waveaudio";
            s += " alias " + al;
            if (shareable) s += " shareable";
            e = send(s + flags, &r, use_notify);
            if (!e) {
                (void)old;                      /* VB6 leaves the device open before open (oracle f_mci) */
                alias = al; devid = r.toInt();
                /* TimeFormat becomes the one of the device (AVIVideo: frames) */
                QString f;
                if (!send("status " + alias + " time format", &f)) {
                    static const char *const names[] = { "milliseconds", "hms", "msf", "frames", "smpte 24", "smpte 25",
                                                          "smpte 30", "smpte 30 drop", "bytes", "samples", "tmsf" };
                    for (int i = 0; i < 11; i++) if (f == names[i]) tf = i;
                }
                caps_valid = false;
                if (hwnddisp) send(QString("window %1 handle %2").arg(alias).arg(hwnddisp));
                if (silent) send("set " + alias + " audio all off");
                restart_updates();
            }
        } else if (!is_open() && cmd != "sound") {
            e = 257;                            /* MCIERR_INVALID_DEVICE_ID: no device */
        } else if (cmd == "close") {
            e = send("close " + alias + flags, &r, use_notify);
            alias.clear(); devid = 0;
            restart_updates();
            if (!e) e = 257;                    /* VB6: no device now (oracle f_mci) */
        } else if (cmd == "play") { e = send("play " + alias + fromto() + flags, &r, use_notify); if (!e && use_notify) play_pending = true; }
        else if ((cmd == "back" || cmd == "step") && !devtype_has_step()) e = 0;     /* VB6: not sent */
        else if (cmd == "eject" && !cap("can eject")) e = 0;
        else if (cmd == "pause" && mode() != 526 && mode() != 527) e = 0;
        else if (cmd == "pause") e = send("pause " + alias + flags, &r, use_notify);
        else if (cmd == "stop") e = send("stop " + alias + flags, &r, use_notify);
        else if (cmd == "back") e = send(QString("step %1 reverse by %2").arg(alias).arg(frames) + flags, &r, use_notify);
        else if (cmd == "step") e = send(QString("step %1 by %2").arg(alias).arg(frames) + flags, &r, use_notify);
        else if (cmd == "prev") e = send("seek " + alias + " to start" + flags, &r, use_notify);
        else if (cmd == "next") e = send("seek " + alias + " to end" + flags, &r, use_notify);
        else if (cmd == "seek") e = send("seek " + alias + " to " + (to_set ? time_text(to) : QString("start")) + flags, &r, use_notify);
        else if (cmd == "record") e = send("record " + alias + (recmode ? " overwrite" : " insert") + fromto() + flags, &r, use_notify);
        else if (cmd == "eject") e = send("set " + alias + " door open" + flags, &r, use_notify);
        else if (cmd == "save") e = send("save " + alias + " \"" + filename + "\"" + flags, &r, use_notify);
        else if (cmd == "sound") {
            QByteArray f = native(filename).toUtf8();
            e = send(QString("play \"%1\"").arg(QString::fromUtf8(f)));
        } else return error;                  /* a command it does not know: nothing (VB6) */
        (void)raise;
        from_set = to_set = false;              /* From / To keep their values, for the next command only */
        set_error(e);
        sync();
        return e;
    }

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        auto vbl = [&](bool &b) -> vbvar { if (set) { b = vb_v2bool(val) != 0; sync(); return vb_v_empty(); } return vbool(b); };
        auto ro = [&]() { vb_raise_full(383, vs(QString("MMControl")), vs(QString("Property is read-only"))); };
        for (int i = 0; i < MB_N; i++) {
            std::string b = mb_name[i];
            if (n == b + "visible") return vbl(vis[i]);
            if (n == b + "enabled") return vbl(en[i]);
        }
        if (n == "devicetype") { if (set) { devtype = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(devtype); }
        if (n == "filename") { if (set) { filename = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(filename); }
        if (n == "command") {
            if (!set) return vstr(command);
            command = qs(vb_v2s(val));
            run(command, true);
            return vb_v_empty();
        }
        if (n == "autoenable") return vbl(autoenable);
        if (n == "notify") { if (set && vb_v2bool(val)) notify_armed = true; return vbl(notify); }
        if (n == "wait") { if (set) wait_armed = vb_v2bool(val) != 0; return vbl(wait); }
        if (n == "shareable") return vbl(shareable);
        if (n == "silent") {
            if (!set) return vbool(silent);
            silent = vb_v2bool(val) != 0;
            if (is_open()) send("set " + alias + (silent ? " audio all off" : " audio all on"));
            return vb_v_empty();
        }
        if (n == "enabled") { if (set) { enabled = vb_v2bool(val) != 0; if (w) w->setEnabled(enabled); sync(); return vb_v_empty(); } return vbool(enabled); }
        if (n == "orientation") {
            if (!set) return vb_v_l(orientation);
            int v = (int)vb_v2l(val);
            if (v < 0 || v > 1) { vb_raise(380); return vb_v_empty(); }
            orientation = v; if (w) w->relayout(); return vb_v_empty();
        }
        if (n == "borderstyle") { if (set) { int v = (int)vb_v2l(val); if (v < 0 || v > 1) { vb_raise(380); return vb_v_empty(); } border = v; return vb_v_empty(); } return vb_v_l(border); }
        if (n == "recordmode") { if (set) { int v = (int)vb_v2l(val); if (v < 0 || v > 1) { vb_raise(380); return vb_v_empty(); } recmode = v; return vb_v_empty(); } return vb_v_l(recmode); }
        if (n == "timeformat") {
            if (!set) return vb_v_l(tf);
            int v = (int)vb_v2l(val);
            if (v < 0 || v > 10) { vb_raise(380); return vb_v_empty(); }
            tf = v;
            if (is_open()) send("set " + alias + " time format " + tf_word(tf));
            return vb_v_empty();
        }
        if (n == "updateinterval") {
            if (!set) return vb_v_l(updint);
            int v = (int)vb_v2l(val);
            if (v < 0) { vb_raise(380); return vb_v_empty(); }
            updint = v; restart_updates(); return vb_v_empty();
        }
        if (n == "from") { if (set) { from = (vb_i32)vb_v2l(val); from_set = true; return vb_v_empty(); } return vb_v_l(from); }
        if (n == "to") { if (set) { to = (vb_i32)vb_v2l(val); to_set = true; return vb_v_empty(); } return vb_v_l(to); }
        if (n == "frames") { if (set) { frames = (vb_i32)vb_v2l(val); return vb_v_empty(); } return vb_v_l(frames); }
        if (n == "mode") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(mode()); }
        if (n == "length") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(time_value(q("length"))); }
        if (n == "position") {
            if (set) { ro(); return vb_v_empty(); }
            return vb_v_l(time_value(q("position")));
        }
        if (n == "start") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(time_value(q("start position"))); }
        if (n == "tracks") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(q("number of tracks").toInt()); }
        if (n == "track") { if (set) { track = (vb_i32)vb_v2l(val); return vb_v_empty(); } return vb_v_l(track); }
        if (n == "tracklength") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(time_value(q(QString("length track %1").arg(track)))); }
        if (n == "trackposition") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(time_value(q(QString("position track %1").arg(track)))); }
        if (n == "canplay") { if (set) { ro(); return vb_v_empty(); } return vbool(cap("can play")); }
        if (n == "canrecord") { if (set) { ro(); return vb_v_empty(); } return vbool(cap("can record")); }
        if (n == "caneject") { if (set) { ro(); return vb_v_empty(); } return vbool(cap("can eject")); }
        if (n == "canstep") { if (set) { ro(); return vb_v_empty(); } return vbool(devtype_has_step()); }
        if (n == "useswindows") { if (set) { ro(); return vb_v_empty(); } return vbool(devtype_has_step()); }
        if (n == "error") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(error); }
        if (n == "errormessage") { if (set) { ro(); return vb_v_empty(); } return vstr(errmsg); }
        if (n == "deviceid") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(devid); }
        if (n == "notifyvalue") { if (set) { ro(); return vb_v_empty(); } return vb_v_l(notifyvalue); }
        if (n == "notifymessage") { if (set) { ro(); return vb_v_empty(); } return vstr(QString()); }
        if (n == "hwnddisplay") {
            if (!set) return vb_v_l((vb_i32)hwnddisp);
            hwnddisp = vb_v2l(val);
            if (is_open()) send(QString("window %1 handle %2").arg(alias).arg(hwnddisp));
            mci_wake();
            return vb_v_empty();
        }
        if (n == "refresh" || n == "aboutbox" || n == "oledrag") return vb_v_empty();
        if (n == "oledropmode" || n == "mousepointer") { if (!set) return vb_v_l(0); return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
};
int OcxMMControl::next_alias = 0;

void MciWidget::relayout()
{
    if (!mm) return;
    int nvis = 0;
    for (int i = 0; i < MB_N; i++) if (mm->vis[i]) nvis++;
    if (!nvis) return;
    bool vert = mm->orientation == 1;
    int total = vert ? height() : width(), k = 0;
    for (int i = 0; i < MB_N; i++) {
        if (!mm->vis[i]) continue;
        int a = total * k / nvis, b = total * (k + 1) / nvis;
        if (vert) btn[i]->setGeometry(0, a, width(), b - a);
        else btn[i]->setGeometry(a, 0, b - a, height());
        k++;
    }
}
static Ocx *new_mmcontrol() { return new OcxMMControl; }
