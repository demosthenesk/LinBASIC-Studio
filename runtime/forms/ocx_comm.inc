/* ocx_comm.inc - the Communications control (MSCOMM32.OCX, MSComm): a serial port with the
 * properties, the OnComm events and the errors of VB6. Linux: termios on /dev/tty*; CommPort n is
 * VBRT_COMn when it is set, otherwise the n-th serial port of the computer (the real /dev/ttyS*,
 * then /dev/ttyUSB*, /dev/ttyACM*). Windows: \\.\COMn. Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
/* ---------------------------------------------------------------- the port (runtime/serial.c) */
extern "C" {
int lbser_open(int n, long long *handle);
void lbser_close(long long h);
int lbser_config(long long h, int baud, char parity, int data, double stop, int hs);
void lbser_lines(long long h, int dtr, int rts);
void lbser_break(long long h, int on);
int lbser_read(long long h, char *buf, int n);
int lbser_write(long long h, const char *buf, int n);
int lbser_outq(long long h);
void lbser_flush(long long h, int in, int out);
int lbser_modem(long long h);
}
struct SerialPort {
    long long h = -1;
    bool is_open() const { return h != -1; }
    int open(int n) { long long x = -1; int e = lbser_open(n, &x); if (!e) h = x; return e; }
    void close() { if (is_open()) lbser_close(h); h = -1; }
    bool config(int baud, char parity, int data, double stop, int hs) { return lbser_config(h, baud, parity, data, stop, hs) != 0; }
    void set_line(bool dtr, bool rts) { lbser_lines(h, dtr, rts); }
    void set_break(bool on) { lbser_break(h, on); }
    int read(char *buf, int n) { return lbser_read(h, buf, n); }
    int write(const char *buf, int n) { return lbser_write(h, buf, n); }
    int outq() { return lbser_outq(h); }
    void flush_in() { lbser_flush(h, 1, 0); }
    void flush_out() { lbser_flush(h, 0, 1); }
    int modem() { return lbser_modem(h); }
    long long id() const { return h; }
};

/* ---------------------------------------------------------------- the control */
struct OcxMSComm : Ocx {
    SerialPort port;
    QTimer *timer = nullptr;
    int commport = 1, inbuf = 1024, outbuf = 512, inputlen = 0, inputmode = 0, rth = 0, sth = 0, hs = 0, event = 0;
    int cdto = 0, ctsto = 0, dsrto = 0, interval = 1000;
    bool dtr = true, rts = false, nulldiscard = false, eofenable = false, brk = false;
    QString settings = "9600,n,8,1", parrep = "?";
    QByteArray in;                              /* the input buffer */
    int lastmodem = 0;
    bool send_armed = false;
    bool eof_open = false;                      /* EOFEnable when the port was opened (the event character of the DCB) */

    ~OcxMSComm() override { port.close(); delete timer; }
    QWidget *make(QWidget *parent) override { (void)parent; return nullptr; }
    void load(const char *const *p) override
    {
        commport = (int)propn(p, "CommPort", 1);
        inbuf = (int)propn(p, "InBufferSize", 1024);
        outbuf = (int)propn(p, "OutBufferSize", 512);
        inputlen = (int)propn(p, "InputLen", 0);
        inputmode = (int)propn(p, "InputMode", 0);
        rth = (int)propn(p, "RThreshold", 0);
        sth = (int)propn(p, "SThreshold", 0);
        hs = (int)propn(p, "Handshaking", 0);
        dtr = propn(p, "DTREnable", 1) != 0;
        rts = propn(p, "RTSEnable", 0) != 0;
        nulldiscard = propn(p, "NullDiscard", 0) != 0;
        eofenable = propn(p, "EOFEnable", 0) != 0;
        cdto = (int)propn(p, "CDTimeout", 0);
        ctsto = (int)propn(p, "CTSTimeout", 0);
        dsrto = (int)propn(p, "DSRTimeout", 0);
        /* the .frm of VB6: Settings in four parts (VB6 oracle d_dev), ParityReplace as a character code */
        int baud = (int)propn(p, "BaudRate", 9600), par = (int)propn(p, "ParitySetting", 0);
        int bits = (int)propn(p, "DataBits", 8), stop = (int)propn(p, "StopBits", 0);
        settings = QString("%1,%2,%3,%4").arg(baud).arg(QChar("noems"[par >= 0 && par < 5 ? par : 0]))
                   .arg(bits).arg(stop == 1 ? "1.5" : stop == 2 ? "2" : "1");
        prop(p, "Settings", &settings);         /* a .frm written by hand */
        QString pr;
        if (prop(p, "ParityReplace", &pr)) {
            bool num; int code = pr.toInt(&num);
            parrep = num ? (code ? QString(QChar(code)) : QString()) : pr;
        }
    }

    /* "9600,n,8,1" -> its parts; a field left out keeps its value (VB6: "9600" sets only the speed);
     * false if it is not valid. *norm: the Settings VB6 gives back */
    static bool parse_settings(const QString &s, const QString &old, int *baud, char *par, int *data, double *stop, QString *norm = nullptr)
    {
        QStringList f = s.split(','), o = old.split(',');
        if (f.size() > 4) return false;
        while (o.size() < 4) o << "";
        QString part[4];
        for (int i = 0; i < 4; i++) {
            QString v = i < f.size() ? f[i].trimmed() : QString();
            part[i] = v.isEmpty() ? o[i].trimmed() : v;
        }
        bool ok;
        int b = part[0].toInt(&ok);
        if (!ok) return false;
        static const int rates[] = { 110, 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 28800, 38400, 56000, 57600, 115200, 128000, 256000 };
        bool known = false;
        for (int r : rates) if (r == b) known = true;
        if (!known) return false;
        QString p = part[1].toLower();
        if (p.size() != 1 || !QString("enoms").contains(p)) return false;
        int d = part[2].toInt(&ok);
        if (!ok || d < 4 || d > 8) return false;
        QString st = part[3];
        double sb = st == "1" ? 1 : st == "1.5" ? 1.5 : st == "2" ? 2 : 0;
        if (!sb) return false;
        if (baud) *baud = b;
        if (par) *par = p.toUpper()[0].toLatin1();
        if (data) *data = d;
        if (stop) *stop = sb;
        if (norm) *norm = QString("%1,%2,%3,%4").arg(b).arg(p).arg(d).arg(st);
        return true;
    }
    bool apply_settings()
    {
        int b; char p; int d; double s;
        if (!parse_settings(settings, settings, &b, &p, &d, &s)) return false;
        return port.config(b, p, d, s, hs);
    }
    void fire(int ev)
    {
        event = ev;
        fire_ctl(c, "oncomm");
    }
    /* what has arrived, and the events */
    void poll(bool events)
    {
        if (!port.is_open()) return;
        char buf[4096];
        int got = 0, n;
        bool eof = false;
        while ((n = port.read(buf, sizeof buf)) > 0) {
            for (int i = 0; i < n; i++) {
                if (nulldiscard && buf[i] == 0) continue;
                if (eof_open && buf[i] == 26) eof = true;
                if (in.size() >= inbuf) { if (events) fire(1008); continue; }
                in.append(buf[i]);
                got++;
            }
        }
        if (!events) return;
        if (eof) fire(7);
        if (got && rth > 0 && in.size() >= rth) fire(2);
        if (send_armed && sth > 0 && port.outq() < sth) { send_armed = false; fire(1); }
        int m = port.modem();
        if (m != lastmodem) {
            int ch = m ^ lastmodem;
            lastmodem = m;
            if (ch & 1) fire(3);
            if (ch & 2) fire(4);
            if (ch & 4) fire(5);
            if ((ch & 8) && !(m & 8)) fire(6);
        }
    }
    vbvar not_open() { vb_raise_full(8018, vs(QString("MSComm")), vs(QString("Operation valid only when the port is open"))); return vb_v_empty(); }
    vbvar readonly() { vb_raise_full(383, vs(QString("MSComm")), vs(QString("Property is read-only"))); return vb_v_empty(); }
    vbvar err(int n, const char *msg) { vb_raise_full(n, vs(QString("MSComm")), vs(QString(msg))); return vb_v_empty(); }

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        auto num = [&](int &f, int lo, int hi) -> vbvar {
            if (!set) return vb_v_i((vb_i16)f);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < lo || v > hi) { vb_raise(380); return vb_v_empty(); }
            f = (int)v; return vb_v_empty();
        };
        auto lng = [&](int &f) -> vbvar {
            if (!set) return vb_v_l(f);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0) { vb_raise(380); return vb_v_empty(); }
            f = (int)v; return vb_v_empty();
        };
        if (n == "portopen") {
            if (!set) return vbool(port.is_open());
            bool want = vb_v2bool(val) != 0;
            if (want == port.is_open()) {
                if (want) return err(8005, "Port already open");
                return err(8012, "The device is not open");
            }
            if (!want) {
                port.close();
                in.clear();
                if (timer) timer->stop();
                return vb_v_empty();
            }
            int e = port.open(commport);
            if (e == 8002) return err(8002, "Invalid port number");
            if (e) return err(8005, "Port already open");
            if (!apply_settings()) { port.close(); return err(8015, "Could not set comm state, there may be one or more invalid communications parameters."); }
            port.set_line(dtr, rts);
            eof_open = eofenable;
            in.clear();
            lastmodem = port.modem();
            if (!timer) {
                timer = new QTimer();
                QObject::connect(timer, &QTimer::timeout, [this]() { poll(true); });
            }
            timer->start(10);
            return vb_v_empty();
        }
        if (n == "commport") {
            if (!set) return vb_v_i((vb_i16)commport);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (port.is_open()) return err(8005, "Port already open");
            if (v < 1 || v > 16) return err(8002, "Invalid port number");
            commport = (int)v; return vb_v_empty();
        }
        if (n == "settings") {
            if (!set) return vstr(settings);
            QString s = qs(vb_v2s(val)), norm;
            if (!parse_settings(s, settings, nullptr, nullptr, nullptr, nullptr, &norm)) { vb_raise(380); return vb_v_empty(); }
            QString old = settings;
            settings = norm;
            if (port.is_open() && !apply_settings()) { settings = old; apply_settings(); return err(8015, "Could not set comm state, there may be one or more invalid communications parameters."); }
            return vb_v_empty();
        }
        if (n == "inbuffersize") {
            if (set && port.is_open()) return err(8000, "Operation not valid while the port is opened");
            return num(inbuf, 0, 32767);
        }
        if (n == "outbuffersize") {
            if (set && port.is_open()) return err(8000, "Operation not valid while the port is opened");
            return num(outbuf, 0, 32767);
        }
        if (n == "inputlen") return num(inputlen, 0, 32767);
        if (n == "inputmode") return num(inputmode, 0, 1);
        if (n == "rthreshold") return num(rth, 0, 32767);
        if (n == "sthreshold") return num(sth, 0, 32767);
        if (n == "handshaking") {
            vbvar r = num(hs, 0, 3);
            if (set && port.is_open() && !vb_errpend) apply_settings();
            return r;
        }
        if (n == "cdtimeout") return lng(cdto);
        if (n == "ctstimeout") return lng(ctsto);
        if (n == "dsrtimeout") return lng(dsrto);
        if (n == "interval") return lng(interval);
        if (n == "dtrenable" || n == "rtsenable") {
            bool &f = n == "dtrenable" ? dtr : rts;
            if (!set) return vbool(f);
            f = vb_v2bool(val) != 0;
            if (port.is_open()) port.set_line(dtr, rts);
            return vb_v_empty();
        }
        if (n == "nulldiscard") { if (!set) return vbool(nulldiscard); nulldiscard = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (n == "eofenable") { if (!set) return vbool(eofenable); eofenable = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (n == "parityreplace") {
            if (!set) return vstr(parrep);
            QString s = qs(vb_v2s(val));
            parrep = s.left(1);
            return vb_v_empty();
        }
        if (n == "commevent") { if (set) return readonly(); return vb_v_i((vb_i16)event); }
        if (n == "commid") { if (set) return readonly(); return vb_v_l(port.is_open() ? (vb_i32)port.id() : -1); }
        if (n == "break") {
            if (!set) return vbool(brk);
            brk = vb_v2bool(val) != 0;
            if (port.is_open()) port.set_break(brk);
            return vb_v_empty();
        }
        if (n == "cdholding" || n == "ctsholding" || n == "dsrholding") {
            if (set) return readonly();
            if (!port.is_open()) return vbool(false);
            int m = port.modem();
            return vbool(m & (n == "ctsholding" ? 1 : n == "dsrholding" ? 2 : 4));
        }
        if (n == "inbuffercount") {
            if (!port.is_open()) return set ? vb_v_empty() : vb_v_i(0);
            poll(false);
            if (!set) return vb_v_i((vb_i16)in.size());
            if (vb_v2l(val) != 0) { vb_raise(380); return vb_v_empty(); }
            in.clear(); port.flush_in();
            return vb_v_empty();
        }
        if (n == "outbuffercount") {
            if (!port.is_open()) return set ? vb_v_empty() : vb_v_i(0);
            if (!set) return vb_v_i((vb_i16)port.outq());
            if (vb_v2l(val) != 0) { vb_raise(380); return vb_v_empty(); }
            port.flush_out();
            return vb_v_empty();
        }
        if (n == "input") {
            if (set) return readonly();
            if (!port.is_open()) return not_open();
            poll(false);
            int k = inputlen > 0 ? std::min(inputlen, (int)in.size()) : (int)in.size();
            QByteArray part = in.left(k);
            in.remove(0, k);
            if (inputmode == 1) {
                vbarr *a = vb_atmp();
                if (part.isEmpty()) { vb_aempty(a, 1, VBA_PLAIN); return vb_v_arr(a, VT_UI1); }   /* Byte(0 To -1) */
                vb_adim(a, 1, VBA_PLAIN, 0, 1, 0, (vb_i32)part.size() - 1);
                if (vb_errpend) return vb_v_empty();
                memcpy(a->data, part.constData(), (size_t)part.size());
                return vb_v_arr(a, VT_UI1);
            }
            vbs s = vb_stmp((vb_i32)part.size());
            for (int i = 0; i < part.size(); i++) s.p[i] = vb_ansi2w((unsigned char)part[i]);
            return vb_v_s(s);
        }
        if (n == "output") {
            if (!set) { vb_raise(394); return vb_v_empty(); }
            if (!port.is_open()) return not_open();
            QByteArray data;
            if (val.vt == VT_BSTR) {
                vbs s = vb_v2s(val);
                for (vb_i32 i = 0; i < s.len; i++) { int ch = vb_w2ansi(s.p[i]); data.append((char)(ch < 0 ? '?' : ch)); }
            } else if (val.vt == (VT_ARRAY | VT_UI1)) {
                vbarr *a = vb_v_arrptr(val);
                if (a && a->nd) data = QByteArray((const char *)a->data, (int)a->n);
            } else { vb_raise(380); return vb_v_empty(); }
            port.write(data.constData(), data.size());
            if (sth > 0) send_armed = true;
            return vb_v_empty();
        }
        if (n == "aboutbox") return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_mscomm() { return new OcxMSComm; }
