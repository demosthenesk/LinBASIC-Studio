/* ocx_anim.inc - Animation (MSCOMCT2.OCX): silent AVI clips, uncompressed or RLE8, as the
 * Windows animation control plays them. Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

/* an AVI clip: its frames decoded */
struct AviClip {
    int usec = 66666;                           /* per frame */
    std::vector<QImage> frames;
};
static quint32 avi_u32(const QByteArray &d, int p) { return p + 4 <= d.size() ? qFromLittleEndian<quint32>(d.constData() + p) : 0; }
static quint16 avi_u16(const QByteArray &d, int p) { return p + 2 <= d.size() ? qFromLittleEndian<quint16>(d.constData() + p) : 0; }

/* one frame of a DIB: BI_RGB (8, 16, 24, 32 bits) or BI_RLE8 drawn over the frame before */
static bool avi_frame(const QByteArray &bits, int w, int h, int bpp, int comp, const QVector<QRgb> &pal, QImage &img)
{
    bool up = h > 0;
    int ah = std::abs(h);
    auto put = [&](int x, int y, QRgb c) { if (x >= 0 && x < w && y >= 0 && y < ah) img.setPixel(x, up ? ah - 1 - y : y, c); };
    if (comp == 1) {                            /* RLE8 */
        int x = 0, y = 0, p = 0, n = bits.size();
        while (p + 1 < n) {
            int a = (uchar)bits[p], b = (uchar)bits[p + 1]; p += 2;
            if (a > 0) { QRgb c = b < pal.size() ? pal[b] : 0; for (int i = 0; i < a; i++) put(x++, y, c); continue; }
            if (b == 0) { x = 0; y++; continue; }
            if (b == 1) break;
            if (b == 2) { if (p + 1 >= n) break; x += (uchar)bits[p]; y += (uchar)bits[p + 1]; p += 2; continue; }
            for (int i = 0; i < b && p < n; i++, p++) { int k = (uchar)bits[p]; put(x++, y, k < pal.size() ? pal[k] : 0); }
            if (b & 1) p++;
        }
        return true;
    }
    if (comp != 0) return false;
    int stride = ((w * bpp + 31) / 32) * 4;
    if (bits.size() < stride * ah) return false;
    for (int y = 0; y < ah; y++) {
        const uchar *row = (const uchar *)bits.constData() + y * stride;
        for (int x = 0; x < w; x++) {
            QRgb c = 0;
            switch (bpp) {
            case 8: { int k = row[x]; c = k < pal.size() ? pal[k] : 0; break; }
            case 16: { int v = row[x * 2] | row[x * 2 + 1] << 8; c = qRgb((v >> 10 & 31) * 255 / 31, (v >> 5 & 31) * 255 / 31, (v & 31) * 255 / 31); break; }
            case 24: c = qRgb(row[x * 3 + 2], row[x * 3 + 1], row[x * 3]); break;
            case 32: c = qRgb(row[x * 4 + 2], row[x * 4 + 1], row[x * 4]); break;
            default: return false;
            }
            put(x, y, c);
        }
    }
    return true;
}
/* RIFF AVI: avih (frame time), strf (BITMAPINFOHEADER + palette), the ##db / ##dc chunks of movi */
static bool avi_load(const QByteArray &d, AviClip *clip)
{
    if (d.size() < 12 || d.left(4) != "RIFF" || d.mid(8, 4) != "AVI ") return false;
    int w = 0, h = 0, bpp = 0, comp = -1;
    QVector<QRgb> pal;
    QImage cur;
    std::function<bool(int, int)> walk = [&](int p, int end) -> bool {
        while (p + 8 <= end) {
            QByteArray id = d.mid(p, 4);
            int sz = (int)avi_u32(d, p + 4), body = p + 8;
            if (sz < 0 || body + sz > d.size()) sz = d.size() - body;
            if (id == "LIST" || id == "RIFF") { if (!walk(body + 4, body + sz)) return false; }
            else if (id == "avih") { clip->usec = (int)avi_u32(d, body); if (clip->usec <= 0) clip->usec = 66666; }
            else if (id == "strf" && comp < 0) {
                int bs = (int)avi_u32(d, body);
                w = (int)(qint32)avi_u32(d, body + 4); h = (int)(qint32)avi_u32(d, body + 8);
                bpp = avi_u16(d, body + 14); comp = (int)avi_u32(d, body + 16);
                int used = (int)avi_u32(d, body + 32);
                if (bpp <= 8 && !used) used = 1 << bpp;
                for (int i = 0; i < used && body + bs + i * 4 + 4 <= body + sz; i++) {
                    int q = body + bs + i * 4;
                    pal.append(qRgb((uchar)d[q + 2], (uchar)d[q + 1], (uchar)d[q]));
                }
                if (w <= 0 || h == 0 || (comp != 0 && comp != 1)) return false;
                cur = QImage(w, std::abs(h), QImage::Format_RGB32);
                cur.fill(pal.isEmpty() ? 0 : pal[0]);
            } else if (id.size() == 4 && id.mid(0, 2) == "00" && (id.mid(2) == "db" || id.mid(2) == "dc")) {
                if (comp < 0) return false;
                if (sz > 0) {
                    if (comp == 0) cur.fill(0);
                    if (!avi_frame(d.mid(body, sz), w, h, bpp, comp, pal, cur)) return false;
                }
                clip->frames.push_back(cur.copy());
            }
            p = body + sz + (sz & 1);
        }
        return true;
    };
    if (!walk(12, d.size())) return false;
    return !clip->frames.empty();
}

class AnimWidget : public QWidget {
public:
    struct OcxAnimation *anim = nullptr;
    explicit AnimWidget(QWidget *parent) : QWidget(parent) {}
protected:
    void paintEvent(QPaintEvent *) override;
};
struct OcxAnimation : Ocx {
    QPointer<AnimWidget> w;
    QTimer *timer = nullptr;
    AviClip clip;
    bool open = false, autoplay = false, center = true;
    int backstyle = 0, frame = 0, first = 0, last = -1, repeat = -1;
    QWidget *make(QWidget *parent) override
    {
        w = new AnimWidget(parent);
        w->anim = this;
        timer = new QTimer(w);
        QObject::connect(timer, &QTimer::timeout, [this]() { step(); });
        return w;
    }
    void load(const char *const *p) override
    {
        autoplay = propn(p, "AutoPlay", 0) != 0;
        center = propn(p, "Center", 0) != 0;
        backstyle = (int)propn(p, "BackStyle", 0);
    }
    void step()
    {
        if (!open || clip.frames.empty()) { timer->stop(); return; }
        int end = last < 0 || last >= (int)clip.frames.size() ? (int)clip.frames.size() - 1 : last;
        if (frame < end) frame++;
        else if (repeat < 0 || --repeat > 0) frame = first;
        else timer->stop();
        if (w) w->update();
    }
    void play(int rep, int from, int to)
    {
        repeat = rep == 0 ? 1 : rep; first = std::max(0, from); last = to;
        frame = std::min(first, (int)clip.frames.size() - 1);
        timer->start(std::max(10, clip.usec / 1000));
        if (w) w->update();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        if (n == "open") {
            if (argc < 1) { vb_raise(450); return vb_v_empty(); }
            QString path = arg_s(argc, argv, 0);
            timer->stop();
            clip = AviClip(); open = false;
            QFile f(native(path));
            if (!f.exists()) { vb_raise(53); return vb_v_empty(); }
            if (!f.open(QIODevice::ReadOnly) || !avi_load(f.readAll(), &clip)) {
                clip = AviClip();
                vb_raise_full(35752, vs(QString("Animation")), vs(QString("Unable to open AVI file")));
                if (w) w->update();
                return vb_v_empty();
            }
            open = true; frame = 0;
            if (autoplay) play(-1, 0, -1);
            if (w) w->update();
            return vb_v_empty();
        }
        if (n == "play") {
            if (!open) { vb_raise_full(35755, vs(QString("Animation")), vs(QString("Video not open"))); return vb_v_empty(); }
            int rep = given(argc, argv, 0) ? (int)vb_v2l(argv[0]) : -1;
            int from = given(argc, argv, 1) ? (int)vb_v2l(argv[1]) : 0;
            int to = given(argc, argv, 2) ? (int)vb_v2l(argv[2]) : -1;
            if (vb_errpend) return vb_v_empty();
            if (from < 0 || from >= (int)clip.frames.size() || (to >= 0 && to < from)) { vb_raise(380); return vb_v_empty(); }
            if (to >= (int)clip.frames.size()) to = -1;            /* past the end: the last frame */
            play(rep, from, to);
            return vb_v_empty();
        }
        if (n == "stop") {
            if (autoplay && open) { vb_raise_full(35759, vs(QString("Animation")), vs(QString("Can't stop AutoPlay animation using the Stop method"))); return vb_v_empty(); }
            timer->stop(); return vb_v_empty();
        }
        if (n == "close") { timer->stop(); clip = AviClip(); open = false; if (w) w->update(); return vb_v_empty(); }
        if (n == "autoplay") {
            if (!set) return vbool(autoplay);
            autoplay = vb_v2bool(val) != 0;
            if (open) { if (autoplay) play(-1, 0, -1); else timer->stop(); }
            return vb_v_empty();
        }
        if (n == "center") { if (set) { vb_raise(382); return vb_v_empty(); } return vbool(center); }   /* design time only */
        if (n == "backstyle") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_l(backstyle); }   /* design time only */
        if (n == "oledropmode") { if (!set) return vb_v_l(0); return vb_v_empty(); }
        if (n == "oledrag" || n == "aboutbox") return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
};
void AnimWidget::paintEvent(QPaintEvent *)
{
    QPainter p(this);
    Ctl *k = anim ? anim->c : nullptr;
    QColor back = k ? qcolor(k->back) : palette().window().color();
    p.fillRect(rect(), back);
    if (!anim || !anim->open || anim->clip.frames.empty()) return;
    QImage img = anim->clip.frames[std::min(anim->frame, (int)anim->clip.frames.size() - 1)];
    if (anim->backstyle == 0) {                 /* transparent: the color of the first pixel is the background */
        QImage t = img.convertToFormat(QImage::Format_ARGB32);
        QRgb key = img.pixel(0, img.height() - 1);
        for (int y = 0; y < t.height(); y++) {
            QRgb *line = (QRgb *)t.scanLine(y);
            for (int x = 0; x < t.width(); x++) if ((line[x] & 0xFFFFFF) == (key & 0xFFFFFF)) line[x] = 0;
        }
        img = t;
    }
    QPoint at = anim->center ? QPoint((width() - img.width()) / 2, (height() - img.height()) / 2) : QPoint(0, 0);
    p.drawImage(at, img);
}
