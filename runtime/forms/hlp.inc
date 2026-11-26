/* hlp.inc - the Windows Help (.hlp, WinHelp 3.1/4) files on Linux (10.1; the user's decision 2026-11-26:
 * a reader of our own, as WinHelp is not on Linux). Windows keeps WinHelp itself, as VB6.
 *
 * Written from the public descriptions of the format (helpfile.txt of M. Winterhoff, helpdeco) and checked
 * against the help files of VB-TESTS/DEV-10/h2 (made by HCW 4.03: no compression, phrases, Zeck) and the
 * PSC help files of work/hlp-corpus (Hall compression, bitmaps, metafiles, macros, secondary windows).
 *
 *  - the file system: a B+ tree of the internal files (|SYSTEM, |TOPIC, |FONT, |CONTEXT, |CTXOMAP,
 *    |TTLBTREE, |KWBTREE, |KWDATA, |Phrases or |PhrIndex/|PhrImage, |bmN);
 *  - |TOPIC: blocks of 4 kB (LZ77), the records (topic header 0x02, text 0x20, table 0x23) linked by their
 *    TOPICPOS; the text of a record by phrases (WinHelp 3.1) or Hall compression (WinHelp 4);
 *  - the TOPICOFFSET of the contexts: (block << 15) + the characters before it in that block;
 *  - the window: as the main window of WinHelp 4 (the menus; Help Topics, Back, Options; the non-scrolling
 *    region of a topic), the Help Topics dialog (Contents of the .cnt file, Index of the K keywords, Find),
 *    pop-ups, jumps (green, underlined), pop-up links (green, dotted), the common macros. */

#include <QTextTable>
#include <QTextFrame>
#include <QTreeWidget>
#include <QHeaderView>
#include <QActionGroup>
#include <functional>
#include <set>
#include <algorithm>
#include <cmath>

/* ---------------------------------------------------------------- reading the bytes */
static inline int hlp_u16(const QByteArray &b, int p) { return p >= 0 && p + 2 <= b.size() ? qFromLittleEndian<quint16>(b.constData() + p) : 0; }
static inline qint32 hlp_i32(const QByteArray &b, int p) { return p >= 0 && p + 4 <= b.size() ? qFromLittleEndian<qint32>(b.constData() + p) : 0; }
struct HlpRd {                                  /* a cursor on a record, with the compressed numbers of WinHelp */
    const uchar *d; int n, p = 0;
    explicit HlpRd(const QByteArray &b, int at = 0) : d((const uchar *)b.constData()), n(b.size()), p(at) {}
    bool end() const { return p >= n; }
    int u8() { return p < n ? d[p++] : (p++, 0); }
    int u16() { int a = u8(); return a | (u8() << 8); }
    int i16() { return (qint16)u16(); }
    qint32 i32() { quint32 a = (quint32)u16(); return (qint32)(a | ((quint32)u16() << 16)); }
    bool odd() const { return p < n && (d[p] & 1); }
    qint32 slong() { if (odd()) return (qint32)((quint32)i32() >> 1) - 0x40000000; return (u16() >> 1) - 0x4000; }
    quint32 ulong() { if (odd()) return (quint32)i32() >> 1; return (quint32)u16() >> 1; }
    int uword() { if (odd()) return u16() >> 1; return u8() >> 1; }
    int sint() { if (odd()) return (u16() >> 1) - 0x4000; return (u8() >> 1) - 0x40; }
    QByteArray bytes(int k) { k = std::max(0, std::min(k, n - p)); QByteArray b((const char *)d + p, k); p += k; return b; }
};
static QByteArray hlp_lz77(const QByteArray &src, int limit)
{
    QByteArray out;
    out.reserve(limit > 0 ? limit : src.size() * 3);
    const uchar *s = (const uchar *)src.constData();
    int n = src.size(), i = 0;
    while (i < n) {
        int flags = s[i++];
        for (int bit = 0; bit < 8 && i < n; bit++) {
            if (flags & (1 << bit)) {
                if (i + 1 >= n) { i = n; break; }
                int v = s[i] | (s[i + 1] << 8);
                i += 2;
                int len = (v >> 12) + 3, dist = (v & 0x0FFF) + 1;
                for (int k = 0; k < len; k++) out.append(dist <= out.size() ? out.at(out.size() - dist) : '\0');
            } else out.append((char)s[i++]);
            if (limit > 0 && out.size() >= limit) { out.truncate(limit); return out; }
        }
    }
    return out;
}
static QByteArray hlp_unrle(const QByteArray &src)
{
    QByteArray out;
    const uchar *s = (const uchar *)src.constData();
    int n = src.size(), i = 0;
    while (i < n) {
        int c = s[i++];
        if (c & 0x80) { int k = std::min(c & 0x7F, n - i); out.append((const char *)s + i, k); i += k; }
        else { if (i < n) out.append(c, (char)s[i]); i++; }
    }
    return out;
}
/* the hash of a context string ("IDH_BUTTON"), as the |CONTEXT tree has it */
static qint32 hlp_hash(const QString &s)
{
    quint32 h = 0;
    for (QChar qc : s) {
        int c = qc.unicode() & 0xFF, x;
        if (c >= 'a' && c <= 'z') c -= 0x20;
        if (c >= 'A' && c <= 'Z') x = c - 'A' + 17;
        else if (c >= '1' && c <= '9') x = c - '0';
        else if (c == '0') x = 10;
        else if (c == '.') x = 12;
        else if (c == '_') x = 13;
        else if (c == '!') x = 11;
        else x = (qint8)(c - 0x30);
        h = h * 43 + (quint32)x;
    }
    return (qint32)h;
}

/* ---------------------------------------------------------------- the file */
struct HlpDoc {
    QString file, title, caption, cntname, codec = "windows-1252", config;
    QByteArray raw;
    std::map<QString, QByteArray> F;            /* the internal files */
    int minor = 0, flags = 0, bsize = 4096;
    bool lz = false, hall = false, ok = false;
    qint32 contents = -1;
    std::vector<QByteArray> phr;
    QByteArray buf;                             /* |TOPIC without its block headers, decompressed */
    std::vector<int> bstart;
    struct Rec { int type; qint32 tpos; quint32 off; QByteArray l1, l2; };
    std::vector<Rec> recs;
    struct Topic { QString title; quint32 off; int r0, r1; qint32 nsr, scroll, bck, fwd; };
    std::vector<Topic> tops;
    std::map<qint32, quint32> ctx;              /* the hash of a context string -> its TOPICOFFSET */
    std::map<quint32, quint32> omap;            /* [MAP]: the context number -> its TOPICOFFSET */
    struct Font { QString face; double pt = 10; int attr = 0, family = 0; QColor fg; bool deffg = true; };
    std::vector<Font> fonts;
    struct Win { QString type, name, caption; int flags = 0, x = 0, y = 0, w = 0, h = 0; QColor rgb, nsr; };
    std::vector<Win> wins;
    std::vector<std::pair<QString, std::vector<quint32>>> kw;
    bool kw_read = false, cnt_read = false, words_read = false;
    SmNode toc;
    std::map<int, QImage> pics;
    std::map<QString, std::vector<int>> words;  /* Find: a word -> the topics */
    std::vector<std::pair<QString, int>> marks; /* Bookmark */
    int unknown = 0;                            /* (the check of VBRT_HLP_CHECK) */
};
static QString hlp_text(HlpDoc *d, const QByteArray &b)
{
    bool ok;
    QString s = decode_as(d->codec.toLatin1().constData(), b, &ok);
    if (s.isEmpty() && !b.isEmpty()) s = QString::fromLatin1(b);
    return s;
}
static QString hlp_cstr(HlpDoc *d, const QByteArray &b, int p = 0)
{
    int e = b.indexOf('\0', p);
    return hlp_text(d, b.mid(p, (e < 0 ? b.size() : e) - p));
}
static QByteArray hlp_internal(const QByteArray &raw, qint32 off)
{
    if (off < 0 || off + 9 > raw.size()) return QByteArray();
    qint32 used = hlp_i32(raw, off + 4);
    if (used < 0 || used > raw.size() - off - 9) used = raw.size() - off - 9;
    return raw.mid(off + 9, used);
}
/* the entries of the leaves of a B+ tree; f(tree, position) returns the position after the entry */
template <class Fn> static void hlp_btree(const QByteArray &t, Fn f)
{
    if (t.size() < 38 || hlp_u16(t, 0) != 0x293B) return;
    int ps = hlp_u16(t, 4), root = hlp_u16(t, 26), total = hlp_u16(t, 30), levels = hlp_u16(t, 32);
    if (ps <= 0) return;
    int page = root;
    for (int l = 1; l < levels && page < total; l++) page = hlp_u16(t, 38 + page * ps + 4);
    std::set<int> seen;
    while (page != 0xFFFF && page < total && !seen.count(page)) {
        seen.insert(page);
        int p = 38 + page * ps, n = hlp_u16(t, p + 2), next = hlp_u16(t, p + 6), q = p + 8;
        for (int i = 0; i < n && q < p + ps && q < t.size(); i++) q = f(t, q);
        page = next;
    }
}
static const char *hlp_charset_codec(int cs)
{
    switch (cs) {
    case 161: return "windows-1253"; case 162: return "windows-1254"; case 238: return "windows-1250";
    case 204: return "windows-1251"; case 177: return "windows-1255"; case 178: return "windows-1256";
    case 186: return "windows-1257"; case 128: return "Shift_JIS"; case 134: return "GBK";
    case 136: return "Big5"; case 129: return "EUC-KR"; case 222: return "windows-874";
    }
    return nullptr;
}
static void hlp_unphrase(HlpDoc *d, const QByteArray &src, int want, QByteArray &out)
{
    const uchar *s = (const uchar *)src.constData();
    int n = src.size(), i = 0;
    const auto &P = d->phr;
    auto put = [&](size_t k) { if (k < P.size()) out += P[k]; };
    while (i < n && out.size() < want) {
        int c = s[i++];
        if (d->hall) {
            if (!(c & 1)) put(c >> 1);
            else if ((c & 3) == 1) { put(128 + (c >> 2) * 256 + (i < n ? s[i] : 0)); i++; }
            else if ((c & 7) == 3) { int k = std::min((c >> 3) + 1, n - i); out.append((const char *)s + i, k); i += k; }
            else if ((c & 0x0F) == 7) out.append((c >> 4) + 1, ' ');
            else out.append((c >> 4) + 1, '\0');
        } else if (c > 0 && c < 0x10 && !P.empty()) {
            int idx = ((c - 1) << 8) | (i < n ? s[i] : 0);
            i++;
            put(idx >> 1);
            if (idx & 1) out.append(' ');
        } else out.append((char)c);
    }
}
static void hlp_phrases(HlpDoc *d)
{
    auto it = d->F.find("|Phrases");
    if (it != d->F.end()) {
        const QByteArray &P = it->second;
        int n = hlp_u16(P, 0);
        if (n <= 0) return;
        int dsize = hlp_i32(P, 4), base = 8 + 2 * (n + 1);
        if (base > P.size()) return;
        QByteArray text = hlp_lz77(P.mid(base), dsize);
        if (text.size() < dsize) text = P.mid(base);
        int o0 = hlp_u16(P, 8);
        for (int k = 0; k < n; k++) {
            int a = hlp_u16(P, 8 + 2 * k) - o0, b = hlp_u16(P, 10 + 2 * k) - o0;
            d->phr.push_back(a >= 0 && b >= a ? text.mid(a, b - a) : QByteArray());
        }
        return;
    }
    auto ii = d->F.find("|PhrIndex"), mi = d->F.find("|PhrImage");
    if (ii == d->F.end() || mi == d->F.end()) return;
    const QByteArray &I = ii->second;
    int entries = hlp_i32(I, 4), imgsize = hlp_i32(I, 12), imgcsize = hlp_i32(I, 16), bits = hlp_u16(I, 24) & 0x0F;
    if (entries <= 0 || entries > 0x100000) return;
    d->hall = true;
    QByteArray img = imgsize != imgcsize ? hlp_lz77(mi->second, imgsize) : mi->second;
    int pos = 28; quint32 mask = 0, val = 0;
    auto bit = [&]() -> bool {
        mask <<= 1;
        if (!mask) { val = (quint32)hlp_i32(I, pos); pos += 4; mask = 1; }
        return (val & mask) != 0;
    };
    int off = 0;
    for (int k = 0; k < entries; k++) {
        int n = 1;
        while (bit() && pos <= I.size() + 4) n += 1 << bits;
        if (bit()) n += 1;
        if (bits > 1 && bit()) n += 2;
        if (bits > 2 && bit()) n += 4;
        if (bits > 3 && bit()) n += 8;
        if (bits > 4 && bit()) n += 16;
        d->phr.push_back(img.mid(off, n));
        off += n;
    }
}
static void hlp_topics(HlpDoc *d)
{
    auto it = d->F.find("|TOPIC");
    if (it == d->F.end()) return;
    const QByteArray &T = it->second;
    int bs = d->bsize;
    for (int p = 0; p < T.size(); p += bs) {   /* the blocks, without their headers, in one stream */
        d->bstart.push_back(d->buf.size());
        QByteArray blk = T.mid(p + 12, bs - 12);
        d->buf += d->lz ? hlp_lz77(blk, 16384 - 12) : blk;
    }
    auto pos = [d](qint32 tp) -> int {
        int blk = (int)((quint32)tp >> 14), off = tp & 0x3FFF;
        if (tp < 0 || blk >= (int)d->bstart.size()) return -1;
        return d->bstart[blk] + off - 12;
    };
    auto block_of = [d](int p) { return (int)(std::upper_bound(d->bstart.begin(), d->bstart.end(), p) - d->bstart.begin()) - 1; };
    int p = pos(hlp_i32(T, 4));
    std::set<int> seen;
    quint32 charoff = 0; int lastblk = -1;
    const QByteArray &B = d->buf;
    while (p >= 0 && p + 21 <= B.size() && !seen.count(p)) {
        seen.insert(p);
        qint32 bsz = hlp_i32(B, p), dl2 = hlp_i32(B, p + 4), nxt = hlp_i32(B, p + 12), dl1 = hlp_i32(B, p + 16);
        int type = (uchar)B[p + 20];
        if (bsz < 21 || dl1 < 21 || dl1 > bsz || p + bsz > B.size()) break;
        int blk = block_of(p);
        if (blk != lastblk) { charoff = (quint32)blk << 15; lastblk = blk; }
        HlpDoc::Rec r;
        r.type = type;
        r.tpos = (blk << 14) + (p - d->bstart[blk] + 12);
        r.off = charoff;
        r.l1 = B.mid(p + 21, dl1 - 21);
        QByteArray l2 = B.mid(p + dl1, bsz - dl1);
        if (dl2 > l2.size() && dl2 < (1 << 24)) { QByteArray o; hlp_unphrase(d, l2, dl2, o); r.l2 = o; }
        else r.l2 = l2;
        if (type == 0x02) {
            HlpDoc::Topic t;
            t.title = hlp_cstr(d, r.l2);
            t.off = charoff;
            t.r0 = (int)d->recs.size();
            t.r1 = t.r0 + 1;
            t.bck = hlp_i32(r.l1, 4); t.fwd = hlp_i32(r.l1, 8);
            t.nsr = hlp_i32(r.l1, 16); t.scroll = hlp_i32(r.l1, 20);
            d->tops.push_back(t);
        } else if (!d->tops.empty()) d->tops.back().r1 = (int)d->recs.size() + 1;
        if (type == 0x20 || type == 0x23) { HlpRd rd(r.l1); rd.slong(); charoff += (quint32)rd.uword(); }
        d->recs.push_back(r);
        int np = pos(nxt);
        if (np <= p) break;
        p = np;
    }
}
static HlpDoc *hlp_doc(const QString &full)
{
    static std::map<QString, HlpDoc *> docs;
    auto it = docs.find(full);
    if (it != docs.end()) return it->second->ok ? it->second : nullptr;
    auto *d = new HlpDoc;
    docs[full] = d;
    d->file = full;
    QFile f(full);
    if (!f.open(QIODevice::ReadOnly)) return nullptr;
    d->raw = f.readAll();
    const QByteArray &R = d->raw;
    if (hlp_i32(R, 0) != 0x00035F3F) return nullptr;
    QByteArray dir = hlp_internal(R, hlp_i32(R, 4));
    hlp_btree(dir, [d, &R](const QByteArray &t, int q) {
        int e = t.indexOf('\0', q);
        if (e < 0) return (int)t.size();
        d->F[QString::fromLatin1(t.mid(q, e - q))] = hlp_internal(R, hlp_i32(t, e + 1));
        return e + 5;
    });
    const QByteArray &S = d->F["|SYSTEM"];
    if (hlp_u16(S, 0) != 0x036C) return nullptr;
    d->minor = hlp_u16(S, 2);
    d->flags = hlp_u16(S, 10);
    if (d->minor <= 16) return nullptr;         /* WinHelp 3.0 (1990): not read */
    d->lz = (d->flags & 0x0C) != 0;
    d->bsize = d->flags == 8 ? 2048 : 4096;
    QByteArray title, caption;
    int lcid = 0, charset = 0;
    for (int p = 12; p + 4 <= S.size();) {     /* the records: type, size, data */
        int t = hlp_u16(S, p), n = hlp_u16(S, p + 2);
        QByteArray data = S.mid(p + 4, n);
        p += 4 + n;
        switch (t) {
        case 1: title = data.left(data.indexOf('\0') < 0 ? data.size() : data.indexOf('\0')); break;
        case 3: if (n >= 4) d->contents = hlp_i32(data, 0); break;
        case 4: d->config += QString::fromLatin1(data.left(data.indexOf('\0') < 0 ? data.size() : data.indexOf('\0'))) + ";"; break;
        case 6: {                               /* a window: flags, type, name, caption, place, colors */
            if (n < 90) break;
            HlpDoc::Win w;
            w.flags = hlp_u16(data, 0);
            auto str = [&](int a, int len) { QByteArray b = data.mid(a, len); int z = b.indexOf('\0'); return z < 0 ? b : b.left(z); };
            w.type = QString::fromLatin1(str(2, 10)); w.name = QString::fromLatin1(str(12, 9)).toLower();
            caption = str(21, 51);
            w.caption = QString::fromLatin1(caption);
            w.x = hlp_u16(data, 72); w.y = hlp_u16(data, 74); w.w = hlp_u16(data, 76); w.h = hlp_u16(data, 78);
            w.rgb = QColor((uchar)data[82], (uchar)data[83], (uchar)data[84]);
            w.nsr = QColor((uchar)data[86], (uchar)data[87], (uchar)data[88]);
            d->wins.push_back(w);
            break;
        }
        case 9: if (n >= 10) lcid = hlp_u16(data, 8); break;
        case 10: d->cntname = QString::fromLatin1(data.left(data.indexOf('\0') < 0 ? data.size() : data.indexOf('\0'))); break;
        case 11: if (n >= 1) charset = (uchar)data[0]; break;
        }
    }
    if (const char *c = hlp_charset_codec(charset)) d->codec = c;
    else {
        static const std::map<int, const char *> cp = { { 0x408, "windows-1253" }, { 0x419, "windows-1251" }, { 0x402, "windows-1251" },
            { 0x405, "windows-1250" }, { 0x415, "windows-1250" }, { 0x40e, "windows-1250" }, { 0x41f, "windows-1254" },
            { 0x40d, "windows-1255" }, { 0x401, "windows-1256" }, { 0x425, "windows-1257" }, { 0x411, "Shift_JIS" },
            { 0x804, "GBK" }, { 0x404, "Big5" }, { 0x412, "EUC-KR" } };
        auto ci = cp.find(lcid);
        if (ci != cp.end()) d->codec = ci->second;
    }
    d->title = hlp_text(d, title);
    for (auto &w : d->wins) if (w.name == "main" && (w.flags & 0x04)) d->caption = w.caption.isEmpty() ? QString() : hlp_text(d, w.caption.toLatin1());
    const QByteArray &FN = d->F["|FONT"];          /* the fonts: face names, then descriptors of 11 bytes */
    int nf = hlp_u16(FN, 0), nd = hlp_u16(FN, 2), fo = hlp_u16(FN, 4), dof = hlp_u16(FN, 6);
    int fl = nf ? (dof - fo) / nf : 0;
    QStringList faces;
    for (int i = 0; i < nf && fl > 0; i++) faces << hlp_cstr(d, FN.mid(fo + i * fl, fl));
    for (int i = 0; i < nd && dof + i * 11 + 11 <= FN.size(); i++) {
        const char *q = FN.constData() + dof + i * 11;
        HlpDoc::Font ft;
        ft.attr = (uchar)q[0]; ft.pt = (uchar)q[1] / 2.0; ft.family = (uchar)q[2];
        int fn = hlp_u16(FN, dof + i * 11 + 3);
        ft.face = fn < faces.size() ? faces[fn] : QString();
        int r = (uchar)q[5], g = (uchar)q[6], b = (uchar)q[7];
        ft.deffg = r == 1 && g == 1 && b == 0;
        ft.fg = QColor(r, g, b);
        d->fonts.push_back(ft);
    }
    hlp_phrases(d);
    hlp_topics(d);
    hlp_btree(d->F["|CONTEXT"], [d](const QByteArray &t, int q) { d->ctx[hlp_i32(t, q)] = (quint32)hlp_i32(t, q + 4); return q + 8; });
    const QByteArray &OM = d->F["|CTXOMAP"];
    for (int i = 0, n = hlp_u16(OM, 0); i < n; i++) d->omap[(quint32)hlp_i32(OM, 2 + 8 * i)] = (quint32)hlp_i32(OM, 6 + 8 * i);
    d->ok = !d->tops.empty();
    return d->ok ? d : nullptr;
}
/* the topic of a TOPICOFFSET: the last that starts at it or before */
static int hlp_topic_at(HlpDoc *d, quint32 off)
{
    int best = -1;
    for (int i = 0; i < (int)d->tops.size(); i++) if (d->tops[i].off <= off && (best < 0 || d->tops[i].off >= d->tops[best].off)) best = i;
    return best;
}
static int hlp_topic_hash(HlpDoc *d, qint32 h, quint32 *off = nullptr)
{
    auto it = d->ctx.find(h);
    if (it == d->ctx.end()) return -1;
    if (off) *off = it->second;
    return hlp_topic_at(d, it->second);
}
static int hlp_topic_id(HlpDoc *d, quint32 id, quint32 *off = nullptr)
{
    auto it = d->omap.find(id);
    if (it == d->omap.end()) return -1;
    if (off) *off = it->second;
    return hlp_topic_at(d, it->second);
}
static int hlp_contents(HlpDoc *d)
{
    int t = d->contents >= 0 ? hlp_topic_at(d, (quint32)d->contents) : -1;
    return t >= 0 ? t : 0;
}
static void hlp_keywords(HlpDoc *d)
{
    if (d->kw_read) return;
    d->kw_read = true;
    const QByteArray &KD = d->F["|KWDATA"];
    hlp_btree(d->F["|KWBTREE"], [d, &KD](const QByteArray &t, int q) {
        int e = t.indexOf('\0', q);
        if (e < 0) return (int)t.size();
        QString k = hlp_text(d, t.mid(q, e - q));
        int cnt = hlp_u16(t, e + 1); qint32 off = hlp_i32(t, e + 3);
        std::vector<quint32> v;
        for (int i = 0; i < cnt && i < 4096; i++) v.push_back((quint32)hlp_i32(KD, off + 4 * i));
        d->kw.emplace_back(k, v);
        return e + 7;
    });
}
/* the .cnt file of the Contents: ":Base", ":Title", "1 Book", "2 Topic=IDH_X[@file][>window]" */
static void hlp_cnt(HlpDoc *d)
{
    if (d->cnt_read) return;
    d->cnt_read = true;
    QFileInfo fi(d->file);
    QStringList tries;
    if (!d->cntname.isEmpty()) tries << fi.absolutePath() + "/" + d->cntname;
    tries << fi.absolutePath() + "/" + fi.completeBaseName() + ".cnt";
    QFile f;
    for (const QString &t : tries) {
        QFileInfo c(native(t));
        if (!c.exists()) {                      /* the name of the file in another case (h1.cnt, H1.CNT) */
            QDir dir(c.absolutePath());
            for (const QString &e : dir.entryList(QDir::Files)) if (e.compare(c.fileName(), Qt::CaseInsensitive) == 0) { c = QFileInfo(dir.filePath(e)); break; }
        }
        f.setFileName(c.absoluteFilePath());
        if (f.open(QIODevice::ReadOnly)) break;
    }
    if (!f.isOpen()) return;
    QString txt = hlp_text(d, f.readAll());
    std::vector<SmNode *> stack{ &d->toc };
    for (QString l : txt.split('\n')) {
        l = l.trimmed();
        if (l.isEmpty() || l.startsWith(':') || l.startsWith(';')) continue;
        int sp = 0;
        while (sp < l.size() && l[sp].isDigit()) sp++;
        if (!sp) continue;
        int lvl = l.left(sp).toInt();
        QString rest = l.mid(sp).trimmed();
        SmNode n;
        int eq = -1;
        for (int i = 0; i < rest.size(); i++) { if (rest[i] == '\\') { i++; continue; } if (rest[i] == '=') { eq = i; break; } }
        n.name = (eq < 0 ? rest : rest.left(eq)).replace("\\=", "=").trimmed();
        if (eq >= 0) n.locals << rest.mid(eq + 1).trimmed();
        while ((int)stack.size() > lvl) stack.pop_back();
        while ((int)stack.size() < lvl) {       /* a level without its book: under the last entry */
            SmNode *b = stack.back();
            if (b->kids.empty()) break;
            stack.push_back(&b->kids.back());
        }
        stack.back()->kids.push_back(n);
        if (eq < 0) stack.push_back(&stack.back()->kids.back());
    }
}

/* ---------------------------------------------------------------- the pictures (|bmN) */
static QImage hlp_dib(const QByteArray &dib)    /* a packed DIB (BITMAPINFO + bits) */
{
    if (dib.size() < 12) return QImage();
    int hs = hlp_i32(dib, 0), bc, ncol, comp = 0;
    if (hs == 12) { bc = hlp_u16(dib, 10); ncol = bc <= 8 ? 1 << bc : 0; ncol *= 3; }
    else {
        bc = hlp_u16(dib, 14); comp = hlp_i32(dib, 16);
        int used = hlp_i32(dib, 32);
        ncol = (used ? used : bc <= 8 ? 1 << bc : 0) * 4;
        if (comp == 3) ncol += 12;
    }
    QByteArray bmp("BM");
    bmp += QByteArray(12, '\0');
    qToLittleEndian<quint32>((quint32)(14 + dib.size()), (uchar *)bmp.data() + 2);
    qToLittleEndian<quint32>((quint32)(14 + hs + ncol), (uchar *)bmp.data() + 10);
    bmp += dib;
    QImage im;
    im.loadFromData(bmp, "BMP");
    return im;
}
static QImage hlp_wmf(const QByteArray &w, int wpx, int hpx);
static QImage hlp_picture(HlpDoc *d, int num)
{
    auto it = d->pics.find(num);
    if (it != d->pics.end()) return it->second;
    QImage im;
    auto fi = d->F.find(QString("|bm%1").arg(num));
    if (fi == d->F.end()) fi = d->F.find(QString("bm%1").arg(num));
    if (fi != d->F.end()) {
        const QByteArray &B = fi->second;
        int n = hlp_u16(B, 2);
        for (int k = 0; k < n && im.isNull(); k++) {   /* the first picture that can be shown */
            int o = hlp_i32(B, 4 + 4 * k);
            if (o <= 0 || o + 2 > B.size()) continue;
            int type = (uchar)B[o], pack = (uchar)B[o + 1];
            HlpRd r(B, o + 2);
            auto unpack = [&](QByteArray raw) {
                if (pack == 2 || pack == 3) raw = hlp_lz77(raw, 0);
                if (pack == 1 || pack == 3) raw = hlp_unrle(raw);
                return raw;
            };
            if (type == 5 || type == 6) {
                r.ulong(); r.ulong();
                int planes = r.uword(), bits = r.uword();
                quint32 wd = r.ulong(), ht = r.ulong(), used = r.ulong();
                r.ulong();
                quint32 csize = r.ulong();
                r.ulong();
                qint32 coff = r.i32(); r.i32();
                (void)planes;
                if (wd == 0 || ht == 0 || wd > 16384 || ht > 16384 || bits == 0 || bits > 32) continue;
                int ncol = type == 6 ? (used ? (int)used : bits <= 8 ? 1 << bits : 0) : 0;
                QByteArray pal = type == 6 ? B.mid(r.p, 4 * ncol) : QByteArray();
                QByteArray px = unpack(B.mid(o + coff, (int)csize));
                int stride = (((int)wd * bits + 31) / 32) * 4;
                if (type == 5) {                /* a DDB: rows of words, from the top; black and white */
                    int ds = (((int)wd * bits + 15) / 16) * 2;
                    QByteArray rows;
                    for (int y = (int)ht - 1; y >= 0; y--) rows += px.mid(y * ds, ds).leftJustified(stride, '\0', true);
                    px = rows;
                    if (bits == 1) { pal = QByteArray("\0\0\0\0\xff\xff\xff\0", 8); ncol = 2; }
                }
                QByteArray dib(40, '\0');
                uchar *h = (uchar *)dib.data();
                qToLittleEndian<quint32>(40, h); qToLittleEndian<qint32>((qint32)wd, h + 4); qToLittleEndian<qint32>((qint32)ht, h + 8);
                qToLittleEndian<quint16>(1, h + 12); qToLittleEndian<quint16>((quint16)bits, h + 14);
                qToLittleEndian<quint32>((quint32)px.size(), h + 20); qToLittleEndian<quint32>((quint32)ncol, h + 32);
                im = hlp_dib(dib + pal + px);
            } else if (type == 8) {
                r.uword();
                int wd = r.u16(), ht = r.u16();
                r.ulong();
                quint32 csize = r.ulong();
                r.ulong();
                qint32 coff = r.i32();
                int wpx = (int)std::lround(wd * 96.0 / 2540), hpx = (int)std::lround(ht * 96.0 / 2540);   /* 0.01 mm */
                im = hlp_wmf(unpack(B.mid(o + coff, (int)csize)), std::max(1, std::min(wpx, 4096)), std::max(1, std::min(hpx, 4096)));
            }
        }
    }
    d->pics[num] = im;
    return im;
}
/* the metafiles of the help files: the records that draw pictures, lines, shapes and text */
static QImage hlp_wmf(const QByteArray &w, int wpx, int hpx)
{
    QImage im(wpx, hpx, QImage::Format_ARGB32_Premultiplied);
    im.fill(Qt::transparent);
    int p = 0;
    if (hlp_i32(w, 0) == (qint32)0x9AC6CDD7) p = 22;
    p += hlp_u16(w, p + 2) * 2;                 /* the header */
    QPainter g(&im);
    g.setRenderHint(QPainter::Antialiasing, false);
    struct Obj { int kind = 0; QPen pen; QBrush brush; QFont font; };
    std::vector<Obj> objs;
    QPen pen(Qt::black); QBrush brush(Qt::white); QFont font;
    QColor text = Qt::black, bk = Qt::white;
    bool opaque = true;
    int ox = 0, oy = 0, ex = wpx, ey = hpx, align = 0;
    QPoint cur;
    std::vector<std::tuple<int, int, int, int, QPen, QBrush, QFont>> saved;
    auto apply = [&]() { g.setWindow(ox, oy, ex ? ex : 1, ey ? ey : 1); g.setViewport(0, 0, wpx, hpx); };
    apply();
    auto color = [&](int at) { return QColor((uchar)w[at], (uchar)w[at + 1], (uchar)w[at + 2]); };
    while (p + 6 <= w.size()) {
        qint32 sz = hlp_i32(w, p);
        int fn = hlp_u16(w, p + 4), a = p + 6;
        if (sz < 3 || p + sz * 2 > w.size()) break;
        auto s16 = [&](int k) { return (int)(qint16)hlp_u16(w, a + 2 * k); };
        switch (fn) {
        case 0x0000: p = w.size(); continue;
        case 0x020B: oy = s16(0); ox = s16(1); apply(); break;
        case 0x020C: ey = s16(0); ex = s16(1); apply(); break;
        case 0x0201: bk = color(a); break;
        case 0x0209: text = color(a); break;
        case 0x0102: opaque = s16(0) == 2; break;
        case 0x012E: align = s16(0); break;
        case 0x001E: saved.emplace_back(ox, oy, ex, ey, pen, brush, font); break;
        case 0x0127:
            if (!saved.empty()) { std::tie(ox, oy, ex, ey, pen, brush, font) = saved.back(); saved.pop_back(); apply(); }
            break;
        case 0x02FA: case 0x02FC: case 0x02FB: case 0x00F7: case 0x01F9: case 0x0142: {
            Obj o;
            if (fn == 0x02FA) {
                o.kind = 1;
                int st = s16(0), wd = s16(1);
                o.pen = (st & 0x0F) == 5 ? QPen(Qt::NoPen) : QPen(color(a + 6), std::max(wd, 0));
                if ((st & 0x0F) == 1) o.pen.setStyle(Qt::DashLine);
                else if ((st & 0x0F) == 2) o.pen.setStyle(Qt::DotLine);
                o.pen.setCosmetic(wd <= 1);
            } else if (fn == 0x02FC) {
                o.kind = 2;
                int st = s16(0);
                o.brush = st == 1 ? QBrush(Qt::NoBrush) : QBrush(color(a + 2));
            } else if (fn == 0x02FB) {
                o.kind = 3;
                QByteArray face = w.mid(a + 18, 32);
                int z = face.indexOf('\0');
                o.font.setFamily(QString::fromLatin1(z < 0 ? face : face.left(z)));
                o.font.setPixelSize(std::max(1, std::abs(s16(0))));
                o.font.setBold(s16(4) >= 600);
                o.font.setItalic(w[a + 10] != 0);
                o.font.setUnderline(w[a + 11] != 0);
            }
            size_t k = 0;
            while (k < objs.size() && objs[k].kind) k++;
            if (k == objs.size()) objs.push_back(o); else objs[k] = o;
            break;
        }
        case 0x012D: {
            int k = s16(0);
            if (k >= 0 && k < (int)objs.size()) {
                if (objs[k].kind == 1) pen = objs[k].pen;
                else if (objs[k].kind == 2) brush = objs[k].brush;
                else if (objs[k].kind == 3) font = objs[k].font;
            }
            break;
        }
        case 0x01F0: { int k = s16(0); if (k >= 0 && k < (int)objs.size()) objs[k].kind = 0; break; }
        case 0x0214: cur = QPoint(s16(1), s16(0)); break;
        case 0x0213: { QPoint to(s16(1), s16(0)); g.setPen(pen); g.drawLine(cur, to); cur = to; break; }
        case 0x041B: case 0x0418: case 0x061C: {
            int k = fn == 0x061C ? 2 : 0;
            QRect r(QPoint(s16(k + 3), s16(k + 2)), QPoint(s16(k + 1), s16(k + 0)));
            g.setPen(pen); g.setBrush(brush);
            if (fn == 0x0418) g.drawEllipse(r);
            else if (fn == 0x061C) g.drawRoundedRect(r, s16(1) / 2.0, s16(0) / 2.0);
            else g.drawRect(r.adjusted(0, 0, -1, -1));
            break;
        }
        case 0x0324: case 0x0325: {
            int n = s16(0);
            QPolygon pg;
            for (int i = 0; i < n && a + 2 + 4 * i + 4 <= w.size(); i++) pg << QPoint(s16(1 + 2 * i), s16(2 + 2 * i));
            g.setPen(pen);
            if (fn == 0x0324) { g.setBrush(brush); g.drawPolygon(pg); } else g.drawPolyline(pg);
            break;
        }
        case 0x0521: case 0x0A32: {
            int x, y, n, sa;
            if (fn == 0x0521) { n = s16(0); sa = a + 2; int k = (n + 1) / 2; y = s16(1 + k); x = s16(2 + k); }
            else { y = s16(0); x = s16(1); n = s16(2); int opt = s16(3); sa = a + 8 + ((opt & 6) ? 8 : 0); }
            QString s = QString::fromLatin1(w.mid(sa, n));
            g.setFont(font); g.setPen(text);
            QFontMetrics fm(font);
            if (align & 6) x -= (align & 6) == 6 ? fm.horizontalAdvance(s) / 2 : fm.horizontalAdvance(s);
            int base = (align & 0x18) == 0x18 ? y : (align & 8) ? y - fm.descent() : y + fm.ascent();
            if (opaque) g.fillRect(QRect(x, base - fm.ascent(), fm.horizontalAdvance(s), fm.height()), bk);
            g.drawText(QPoint(x, base), s);
            break;
        }
        case 0x0F43: case 0x0B41: case 0x0940: {
            int sx, sy, sw, sh, dx, dy, dw, dh, dibat;
            if (fn == 0x0F43) { sh = s16(3); sw = s16(4); sy = s16(5); sx = s16(6); dh = s16(7); dw = s16(8); dy = s16(9); dx = s16(10); dibat = a + 22; }
            else if (fn == 0x0B41) { sh = s16(2); sw = s16(3); sy = s16(4); sx = s16(5); dh = s16(6); dw = s16(7); dy = s16(8); dx = s16(9); dibat = a + 20; }
            else { sy = s16(2); sx = s16(3); sh = dh = s16(4); sw = dw = s16(5); dy = s16(6); dx = s16(7); dibat = a + 16; }
            if (dibat >= p + sz * 2) break;     /* (a pattern fill: no bitmap) */
            QImage b = hlp_dib(w.mid(dibat, p + sz * 2 - dibat));
            if (!b.isNull()) g.drawImage(QRect(dx, dy, dw, dh), b, QRect(sx, b.height() - sy - sh, sw, sh).normalized());
            break;
        }
        }
        p += sz * 2;
    }
    g.end();
    return im;
}

/* ---------------------------------------------------------------- a topic -> a QTextDocument */
struct HlpLink { QString href; };
struct HlpOut {                                 /* what the window needs after the rendering */
    std::vector<std::pair<quint32, int>> marks; /* the TOPICOFFSET of a record -> its position in the scrolling part */
    QStringList macros;
};
static double hlp_scale = 1.0;                  /* Options > Font: Small, Normal, Large */
static bool hlp_syscolors = false;              /* Options > Use System Colors */
static int hlp_tw(int v) { return (int)std::lround(v * 10 / 15.0); }   /* 10 twips -> pixels (96 dpi) */
static QTextCharFormat hlp_font(HlpDoc *d, int k)
{
    QTextCharFormat f;
    if (k < 0 || k >= (int)d->fonts.size()) { f.setFontPointSize(10 * hlp_scale); return f; }
    const HlpDoc::Font &ft = d->fonts[k];
    QString face = ft.face;
    if (face.compare("Helv", Qt::CaseInsensitive) == 0 || face.compare("MS Sans Serif", Qt::CaseInsensitive) == 0) face = "Arial";
    else if (face.compare("Tms Rmn", Qt::CaseInsensitive) == 0 || face.compare("MS Serif", Qt::CaseInsensitive) == 0) face = "Times New Roman";
#if QT_VERSION >= QT_VERSION_CHECK(6, 0, 0)
    f.setFontFamilies({ face });
#else
    f.setFontFamily(face);
#endif
    f.setFontStyleHint(ft.family == 1 ? QFont::TypeWriter : ft.family == 2 ? QFont::Serif : ft.family == 4 ? QFont::Cursive :
                       ft.family == 5 ? QFont::Decorative : QFont::SansSerif);
    if (ft.family == 1) f.setFontFixedPitch(true);
    f.setFontPointSize(std::max(1.0, ft.pt) * hlp_scale);
    f.setFontWeight(ft.attr & 1 ? QFont::Bold : QFont::Normal);
    f.setFontItalic(ft.attr & 2);
    f.setFontUnderline(ft.attr & 0x14);
    f.setFontStrikeOut(ft.attr & 8);
    if (ft.attr & 0x20) f.setFontCapitalization(QFont::SmallCaps);
    if (!ft.deffg && !hlp_syscolors) f.setForeground(ft.fg);
    return f;
}
static void hlp_render(HlpDoc *d, int ti, QTextDocument *scroll, QTextDocument *nsr, HlpOut *out)
{
    const HlpDoc::Topic &t = d->tops[ti];
    QTextCharFormat base = hlp_font(d, 0);
    scroll->setDocumentMargin(8);
    if (nsr) nsr->setDocumentMargin(8);
    QTextCursor cs(scroll), cn(nsr ? nsr : scroll);
    bool fresh_s = true, fresh_n = true;
    for (int ri = t.r0 + 1; ri < t.r1 && ri < (int)d->recs.size(); ri++) {
        const HlpDoc::Rec &r = d->recs[ri];
        if (r.type != 0x20 && r.type != 0x23) continue;
        bool in_nsr = nsr && t.nsr != -1 && r.tpos >= t.nsr && (t.scroll == -1 || r.tpos < t.scroll);
        QTextCursor &c = in_nsr ? cn : cs;
        bool &fresh = in_nsr ? fresh_n : fresh_s;
        QTextDocument *doc = c.document();
        if (!in_nsr && out) out->marks.emplace_back(r.off, c.position());
        QList<QByteArray> texts = r.l2.split('\0');
        int ti2 = 0;
        HlpRd rd(r.l1);
        rd.slong(); rd.uword();
        QTextTable *table = nullptr;
        std::vector<int> colw;
        if (r.type == 0x23) {                   /* a table: one row of this record */
            int ncol = rd.u8(), ttype = rd.u8();
            if (ttype == 0 || ttype == 2) rd.i16();
            for (int k = 0; k < ncol; k++) { int wd = rd.i16(); rd.i16(); colw.push_back(wd); }
            if (ncol <= 0 || ncol > 64) { d->unknown++; continue; }
            if (!fresh) c.insertBlock();
            fresh = false;
            QTextTableFormat tf;
            tf.setBorder(0);
            tf.setCellPadding(2);
            tf.setCellSpacing(0);
            QVector<QTextLength> cw;
            for (int wd : colw) cw << QTextLength(QTextLength::FixedLength, std::max(8, hlp_tw(wd)));
            tf.setColumnWidthConstraints(cw);
            table = c.insertTable(1, ncol, tf);
        }
        QTextCharFormat fmt = base;
        QString href; bool styled = false, popup = false;
        bool pend = false;                      /* a paragraph ended: a new block before the next text */
        QTextBlockFormat bf;
        auto charfmt = [&]() {
            QTextCharFormat f = fmt;
            if (!href.isEmpty()) {
                f.setAnchor(true);
                f.setAnchorHref(href);
                if (styled) {
                    f.setForeground(hlp_syscolors ? QApplication::palette().link().color() : QColor(0, 0x80, 0));
                    f.setFontUnderline(true);
                    f.setUnderlineStyle(popup ? QTextCharFormat::DotLine : QTextCharFormat::SingleUnderline);
                }
            }
            return f;
        };
        auto block = [&]() {
            if (pend) { c.insertBlock(bf); pend = false; }
            else if (fresh) { c.setBlockFormat(bf); }
            fresh = false;
        };
        auto put = [&](const QString &s) { if (s.isEmpty()) return; block(); c.insertText(s, charfmt()); };
        bool done = false;
        while (!done) {
            if (table) {
                int col = rd.i16();
                if (col == -1 || rd.end()) break;
                rd.i16(); rd.u8();
                if (col >= 0 && col < table->columns()) c = table->cellAt(0, col).firstCursorPosition();
                pend = false; fresh = true;
            }
            rd.u8(); rd.u8(); rd.u16();         /* (unknown), id */
            int bits = rd.u16();
            bf = QTextBlockFormat();
            int above = 0, below = 0, lines = 0, left = 0, right = 0, first = 0;
            if (bits & 0x0001) rd.slong();
            if (bits & 0x0002) above = rd.sint();
            if (bits & 0x0004) below = rd.sint();
            if (bits & 0x0008) lines = rd.sint();
            if (bits & 0x0010) left = rd.sint();
            if (bits & 0x0020) right = rd.sint();
            if (bits & 0x0040) first = rd.sint();
            if (bits & 0x0100) { rd.u8(); rd.i16(); }
            QList<QTextOption::Tab> tabs;
            if (bits & 0x0200) {
                int nt = rd.sint();
                for (int k = 0; k < nt && k < 64; k++) {
                    int ts = rd.uword(), tt = 0;
                    if (ts & 0x4000) tt = rd.uword();
                    tabs << QTextOption::Tab(hlp_tw(ts & 0x3FFF), tt == 1 ? QTextOption::RightTab : tt == 2 ? QTextOption::CenterTab : QTextOption::LeftTab);
                }
            }
            bf.setTopMargin(hlp_tw(above));
            bf.setBottomMargin(hlp_tw(below));
            bf.setLeftMargin(hlp_tw(left));
            bf.setRightMargin(hlp_tw(right));
            bf.setTextIndent(hlp_tw(first));
            if (lines > 0) bf.setLineHeight(hlp_tw(lines), QTextBlockFormat::MinimumHeight);
            if (!tabs.isEmpty()) bf.setTabPositions(tabs);
            bf.setAlignment(bits & 0x0400 ? Qt::AlignRight : bits & 0x0800 ? Qt::AlignHCenter : Qt::AlignLeft);
            if (!fresh && !pend) pend = true;
            for (;;) {
                QByteArray s = ti2 < texts.size() ? texts[ti2] : QByteArray();
                ti2++;
                put(hlp_text(d, s));
                if (rd.end()) { done = true; break; }
                int cmd = rd.u8();
                if (cmd == 0xFF) { if (!table) done = true; break; }
                switch (cmd) {
                case 0x20: rd.i32(); break;
                case 0x21: rd.i16(); break;
                case 0x80: fmt = hlp_font(d, rd.u16()); break;
                case 0x81: put(QString(QChar(0x2028))); break;
                case 0x82: block(); pend = true; break;
                case 0x83: put("\t"); break;
                case 0x86: case 0x87: case 0x88: {
                    int kind = rd.u8();
                    qint32 psize = rd.slong();
                    if (kind == 0x22) rd.uword();
                    int at = rd.p;
                    int ptype = rd.i16(), num = rd.u16();
                    rd.p = at + std::max(0, (int)psize);
                    if (kind != 0x22 && kind != 0x03) break;
                    QImage im = ptype == 0 ? hlp_picture(d, num) : QImage();
                    if (im.isNull()) break;
                    block();
                    QString name = QString("hlpbm:%1").arg(num);
                    doc->addResource(QTextDocument::ImageResource, QUrl(name), im);
                    QTextImageFormat imf;
                    QTextCharFormat cf = charfmt();
                    imf.merge(cf);
                    imf.setName(name);
                    imf.setWidth(im.width()); imf.setHeight(im.height());
                    if (cmd == 0x86) c.insertImage(imf);
                    else c.insertImage(imf, cmd == 0x87 ? QTextFrameFormat::FloatLeft : QTextFrameFormat::FloatRight);
                    break;
                }
                case 0x89: href.clear(); break;
                case 0x8B: put(QString(QChar(0x00A0))); break;
                case 0x8C: put(QString(QChar(0x2011))); break;
                case 0xC8: case 0xCC: {
                    int n = rd.i16();
                    QByteArray m = rd.bytes(n);
                    int z = m.indexOf('\0');
                    if (out) { out->macros << hlp_text(d, z < 0 ? m : m.left(z)); href = QString("hlp:m%1").arg(out->macros.size() - 1); }
                    styled = cmd == 0xC8; popup = false;
                    break;
                }
                case 0xE0: case 0xE1: case 0xE2: case 0xE3: case 0xE6: case 0xE7: {
                    qint32 v = rd.i32();
                    popup = cmd == 0xE0 || cmd == 0xE2 || cmd == 0xE6;
                    styled = cmd < 0xE6;
                    href = QString("hlp:%1%2%3").arg(popup ? 'p' : 'j').arg(cmd <= 0xE1 ? 'o' : 'h').arg((quint32)v);
                    break;
                }
                case 0xEA: case 0xEB: case 0xEE: case 0xEF: {
                    int n = rd.i16();
                    QByteArray x = rd.bytes(n);
                    HlpRd xr(x);
                    int type = xr.u8();
                    quint32 v = (quint32)xr.i32();
                    QString fname;
                    if (type == 1) xr.u8();
                    else if (type == 4 || type == 6) {
                        if (type == 6) { int z = x.indexOf('\0', xr.p); xr.p = z < 0 ? x.size() : z + 1; }
                        fname = hlp_cstr(d, x, xr.p);
                    }
                    popup = cmd == 0xEA || cmd == 0xEE;
                    styled = cmd <= 0xEB;
                    href = QString("hlp:%1h%2").arg(popup ? 'p' : 'j').arg(v) + (fname.isEmpty() ? QString() : "@" + QString::fromLatin1(QUrl::toPercentEncoding(fname)));
                    break;
                }
                default: d->unknown++; done = true; break;
                }
                if (done) break;
            }
            if (!table) break;
        }
        if (table) { c = QTextCursor(doc); c.movePosition(QTextCursor::End); fresh = false; }
    }
}
/* the text of a topic, without its formatting (Find, the log of the tests) */
static QString hlp_plain(HlpDoc *d, int ti)
{
    QTextDocument doc;
    hlp_render(d, ti, &doc, nullptr, nullptr);
    return doc.toPlainText();
}
static QString hlp_topic_name(HlpDoc *d, int ti)
{
    if (ti < 0 || ti >= (int)d->tops.size()) return "-";
    QString s = d->tops[ti].title.trimmed();
    if (s.isEmpty()) s = hlp_plain(d, ti).simplified().left(60);
    return s.isEmpty() ? QString("#%1").arg(d->tops[ti].off) : s;
}

/* ---------------------------------------------------------------- the windows */
static void hlp_go(HlpDoc *d, int ti, quint32 off);
static void hlp_show_popup(HlpDoc *d, int ti, QPoint pt);
static void hlp_run_macros(HlpDoc *d, const QString &macros, QWidget *from);
static void hlp_topics_dlg(HlpDoc *d, int tab, const QString &key);
static void hlp_missing_topic(QWidget *from)
{
    QMessageBox::warning(from, "Windows Help", greek_or("The topic does not exist. Contact your application vendor for an updated Help file.",
                                                        "Το θέμα δεν υπάρχει. Επικοινωνήστε με τον κατασκευαστή της εφαρμογής για ενημερωμένο αρχείο βοήθειας."));
}
/* another help file of a link or a macro: next to this one (any case of its name) */
static QString hlp_resolve(HlpDoc *d, QString name)
{
    name = name.section('>', 0, 0).trimmed();
    if (name.isEmpty()) return d->file;
    QFileInfo fi(native(name));
    if (fi.isRelative()) fi = QFileInfo(QFileInfo(d->file).absolutePath() + "/" + native(name));
    if (!fi.exists()) {
        QDir dir(fi.absolutePath());
        for (const QString &e : dir.entryList(QDir::Files)) if (e.compare(fi.fileName(), Qt::CaseInsensitive) == 0) return dir.filePath(e);
    }
    return fi.absoluteFilePath();
}
static HlpDoc *hlp_other(HlpDoc *d, const QString &file)
{
    QString f = file.section('>', 0, 0).trimmed();
    return f.isEmpty() ? d : hlp_doc(hlp_resolve(d, f));
}
/* a link of a topic: "hlp:jh<hash>", "hlp:ph<hash>", "hlp:jo<offset>", "hlp:m<n>"; "...@<file>" */
static void hlp_follow(HlpDoc *d, const QString &href, const QStringList &macros, QWidget *from, QPoint at)
{
    if (!href.startsWith("hlp:")) { QDesktopServices::openUrl(QUrl(href)); return; }
    QString s = href.mid(4);
    if (s.startsWith('m')) {
        int k = s.mid(1).toInt();
        if (k >= 0 && k < macros.size()) hlp_run_macros(d, macros[k], from);
        return;
    }
    bool popup = s.startsWith('p'), byoff = s.mid(1, 1) == "o";
    QString v = s.mid(2), file;
    int a = v.indexOf('@');
    if (a >= 0) { file = QUrl::fromPercentEncoding(v.mid(a + 1).toLatin1()); v = v.left(a); }
    HlpDoc *t = file.isEmpty() ? d : hlp_other(d, file);
    if (!t) return;
    quint32 off = 0;
    int ti;
    if (byoff) { off = v.toUInt(); ti = hlp_topic_at(t, off); }
    else ti = hlp_topic_hash(t, (qint32)v.toUInt(), &off);
    if (ti < 0) { hlp_missing_topic(from); return; }
    if (popup) hlp_show_popup(t, ti, at);
    else hlp_go(t, ti, off);
}
class HlpBrowser : public QTextBrowser {
public:
    HlpDoc *d;
    HlpOut out;
    explicit HlpBrowser(HlpDoc *doc, QWidget *parent) : QTextBrowser(parent), d(doc)
    {
        setOpenLinks(false);
        setOpenExternalLinks(false);
        setFrameShape(QFrame::NoFrame);
        QObject::connect(this, &QTextBrowser::anchorClicked, [this](const QUrl &u) { link(u.toString(QUrl::FullyEncoded)); });
    }
    std::function<void(const QString &)> link_fn;
    void link(const QString &h) { if (link_fn) link_fn(h); else hlp_follow(d, h, out.macros, window(), QCursor::pos()); }
    void colors(QColor bg)
    {
        QPalette pal = palette();
        pal.setColor(QPalette::Base, bg);
        pal.setColor(QPalette::Text, hlp_syscolors ? QApplication::palette().text().color() : QColor(Qt::black));
        setPalette(pal);
        viewport()->setPalette(pal);
        viewport()->setAutoFillBackground(true);
    }
};

/* ---------------------------------------------------------------- the pop-up of a topic */
class HlpPopupWin : public QFrame {
public:
    HlpBrowser *b;
    HlpPopupWin(HlpDoc *d, int ti, QPoint pt) : QFrame(nullptr, Qt::Popup)
    {
        setAttribute(Qt::WA_DeleteOnClose);
        setFrameStyle(QFrame::Box | QFrame::Plain);
        setLineWidth(1);
        b = new HlpBrowser(d, this);
        b->colors(hlp_syscolors ? QApplication::palette().toolTipBase().color() : QColor(0xFF, 0xFF, 0xE1));
        hlp_render(d, ti, b->document(), nullptr, &b->out);
        b->link_fn = [this, d](const QString &h) {
            QStringList m = b->out.macros;
            QPoint at = QCursor::pos();
            close();
            hlp_follow(d, h, m, nullptr, at);
        };
        b->viewport()->installEventFilter(this);
        QColor bg = hlp_syscolors ? QApplication::palette().toolTipBase().color() : QColor(0xFF, 0xFF, 0xE1);
        QPalette pal = palette();
        pal.setColor(QPalette::Window, bg);
        pal.setColor(QPalette::WindowText, Qt::black);
        setPalette(pal);
        setAutoFillBackground(true);
        auto *lay = new QVBoxLayout(this);
        lay->setContentsMargins(1, 1, 1, 1);
        lay->setSpacing(0);
        lay->addWidget(b);
        QTextDocument tmp;                      /* the size: of a copy (the document of the browser follows its width) */
        hlp_render(d, ti, &tmp, nullptr, nullptr);
        tmp.setTextWidth(4000);
        double iw = tmp.idealWidth();
        int lh = QFontMetrics(b->font()).height();
        int w = std::max(120, std::min((int)std::ceil(iw) + 2, std::max(320, (int)std::sqrt(iw * lh * 8.0))));
        w = std::min(w, 560);
        tmp.setTextWidth(w);
        QRect scr = QGuiApplication::screenAt(pt) ? QGuiApplication::screenAt(pt)->availableGeometry() : QRect(0, 0, 1024, 768);
        int h = (int)std::ceil(tmp.size().height()) + 2;
        if (h > scr.height() * 2 / 3) { h = scr.height() * 2 / 3; w += b->verticalScrollBar()->sizeHint().width(); }
        else b->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
        b->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
        b->setFixedSize(w, h);
        setFixedSize(w + 2, h + 2);
        QPoint at(pt.x() - width() / 2, pt.y());
        at.setX(std::max(scr.left(), std::min(at.x(), scr.right() - width())));
        if (at.y() + height() > scr.bottom()) at.setY(std::max(scr.top(), pt.y() - height()));
        move(at);
    }
protected:
    bool eventFilter(QObject *o, QEvent *e) override
    {
        if (e->type() == QEvent::MouseButtonRelease && b->anchorAt(static_cast<QMouseEvent *>(e)->pos()).isEmpty()) { close(); return true; }
        return QFrame::eventFilter(o, e);
    }
    void keyPressEvent(QKeyEvent *e) override { if (e->matches(QKeySequence::Copy)) b->copy(); else close(); }
};
static void hlp_show_popup(HlpDoc *d, int ti, QPoint pt)
{
    auto *p = new HlpPopupWin(d, ti, pt);
    p->show();
}

/* ---------------------------------------------------------------- the icons of Help Topics (ours) */
static QIcon hlp_icon(int kind)                 /* 0 a closed book, 1 an open book, 2 a page */
{
    QPixmap pm(16, 16);
    pm.fill(Qt::transparent);
    QPainter g(&pm);
    g.setRenderHint(QPainter::Antialiasing, false);
    if (kind == 0) {
        g.fillRect(3, 2, 10, 12, QColor(0x20, 0x40, 0x90));
        g.fillRect(3, 2, 2, 12, QColor(0x10, 0x20, 0x60));
        g.fillRect(12, 3, 2, 11, QColor(0xF0, 0xF0, 0xF0));
        g.setPen(QColor(0x10, 0x20, 0x60)); g.drawRect(3, 2, 10, 12);
    } else if (kind == 1) {
        g.setPen(QColor(0x30, 0x30, 0x30));
        g.setBrush(QColor(0xFF, 0xFF, 0xF4));
        g.drawPolygon(QPolygon({ QPoint(1, 4), QPoint(7, 6), QPoint(7, 14), QPoint(1, 12) }));
        g.drawPolygon(QPolygon({ QPoint(14, 4), QPoint(8, 6), QPoint(8, 14), QPoint(14, 12) }));
        g.setPen(QColor(0x20, 0x40, 0x90)); g.drawLine(2, 13, 7, 15); g.drawLine(13, 13, 8, 15);
    } else {
        g.setPen(QColor(0x50, 0x50, 0x50));
        g.setBrush(Qt::white);
        g.drawRect(3, 1, 10, 13);
        QFont f = g.font(); f.setPixelSize(11); f.setBold(true); g.setFont(f);
        g.setPen(QColor(0x80, 0x00, 0x80));
        g.drawText(QRect(3, 1, 11, 14), Qt::AlignCenter, "?");
    }
    g.end();
    return QIcon(pm);
}

/* ---------------------------------------------------------------- Help Topics: Contents, Index, Find */
class HlpTopics : public QDialog {
public:
    HlpDoc *d;
    QTabWidget *tabs;
    QTreeWidget *toc = nullptr;
    QLineEdit *kw, *fw;
    QListWidget *idx, *fwords, *ftops;
    QLabel *fcount;
    QPushButton *b_main, *b_print;
    int t_toc = -1, t_idx, t_find;
    explicit HlpTopics(HlpDoc *doc) : QDialog(nullptr), d(doc)
    {
        setAttribute(Qt::WA_QuitOnClose, false);
        setWindowTitle(greek_or("Help Topics: ", "Θέματα βοήθειας: ") + (d->title.isEmpty() ? QFileInfo(d->file).completeBaseName() : d->title));
        hlp_cnt(d);
        hlp_keywords(d);
        auto *lay = new QVBoxLayout(this);
        tabs = new QTabWidget;
        lay->addWidget(tabs, 1);
        if (!d->toc.kids.empty()) {             /* Contents: only with a .cnt file, as WinHelp */
            auto *w = new QWidget;
            auto *l = new QVBoxLayout(w);
            l->addWidget(new QLabel(greek_or("Click a book, and then click Open. Or click another tab, such as Index.",
                                             "Επιλέξτε ένα βιβλίο και πατήστε Άνοιγμα. Ή επιλέξτε άλλη καρτέλα, π.χ. Ευρετήριο.")));
            toc = new QTreeWidget;
            toc->setHeaderHidden(true);
            toc->setRootIsDecorated(false);
            toc->setIndentation(18);
            l->addWidget(toc, 1);
            fill(&d->toc, nullptr);
            t_toc = tabs->addTab(w, greek_or("Contents", "Περιεχόμενα"));
            QObject::connect(toc, &QTreeWidget::currentItemChanged, [this]() { buttons(); });
            QObject::connect(toc, &QTreeWidget::itemExpanded, [](QTreeWidgetItem *it) { it->setIcon(0, hlp_icon(1)); });
            QObject::connect(toc, &QTreeWidget::itemCollapsed, [](QTreeWidgetItem *it) { it->setIcon(0, hlp_icon(0)); });
            QObject::connect(toc, &QTreeWidget::itemActivated, [this](QTreeWidgetItem *it) { open_item(it); });
        }
        {
            auto *w = new QWidget;
            auto *l = new QVBoxLayout(w);
            auto *l1 = new QLabel(greek_or("&1  Type the first few letters of the word you're looking for.",
                                           "&1  Πληκτρολογήστε τα πρώτα γράμματα της λέξης που αναζητάτε."));
            l->addWidget(l1);
            kw = new QLineEdit;
            l1->setBuddy(kw);
            l->addWidget(kw);
            auto *l2 = new QLabel(greek_or("&2  Click the index entry you want, and then click Display.",
                                           "&2  Επιλέξτε την καταχώρηση του ευρετηρίου και πατήστε Εμφάνιση."));
            l->addWidget(l2);
            idx = new QListWidget;
            l2->setBuddy(idx);
            l->addWidget(idx, 1);
            for (const auto &k : d->kw) idx->addItem(k.first);
            t_idx = tabs->addTab(w, greek_or("Index", "Ευρετήριο"));
            QObject::connect(kw, &QLineEdit::textEdited, [this](const QString &s) { find_key(s); });
            QObject::connect(kw, &QLineEdit::returnPressed, [this]() { show_key(idx->currentRow()); });
            QObject::connect(idx, &QListWidget::itemActivated, [this](QListWidgetItem *) { show_key(idx->currentRow()); });
        }
        {
            auto *w = new QWidget;
            auto *l = new QVBoxLayout(w);
            auto *f1 = new QLabel(greek_or("&1  Type the word(s) you want to find", "&1  Πληκτρολογήστε τις λέξεις που θέλετε να βρείτε"));
            l->addWidget(f1);
            fw = new QLineEdit;
            f1->setBuddy(fw);
            l->addWidget(fw);
            auto *f2 = new QLabel(greek_or("&2  Select some matching words to narrow your search", "&2  Επιλέξτε λέξεις για να περιορίσετε την αναζήτηση"));
            l->addWidget(f2);
            fwords = new QListWidget;
            f2->setBuddy(fwords);
            fwords->setSelectionMode(QAbstractItemView::ExtendedSelection);
            l->addWidget(fwords, 1);
            auto *f3 = new QLabel(greek_or("&3  Click a topic, then click Display", "&3  Επιλέξτε ένα θέμα και πατήστε Εμφάνιση"));
            l->addWidget(f3);
            ftops = new QListWidget;
            f3->setBuddy(ftops);
            l->addWidget(ftops, 1);
            fcount = new QLabel;
            l->addWidget(fcount);
            t_find = tabs->addTab(w, greek_or("Find", "Εύρεση"));
            QObject::connect(fw, &QLineEdit::textEdited, [this]() { find_words(); });
            QObject::connect(fwords, &QListWidget::itemSelectionChanged, [this]() { find_topics(); });
            QObject::connect(ftops, &QListWidget::itemActivated, [this](QListWidgetItem *it) { go_topic(it->data(Qt::UserRole).toInt()); });
        }
        auto *br = new QHBoxLayout;
        br->addStretch();
        b_main = new QPushButton(greek_or("&Display", "&Εμφάνιση"));
        b_main->setDefault(true);
        b_print = new QPushButton(greek_or("&Print...", "Ε&κτύπωση..."));
        auto *b_cancel = new QPushButton(greek_or("Cancel", "Άκυρο"));
        br->addWidget(b_main); br->addWidget(b_print); br->addWidget(b_cancel);
        lay->addLayout(br);
        QObject::connect(b_cancel, &QPushButton::clicked, this, &QDialog::reject);
        QObject::connect(b_main, &QPushButton::clicked, [this]() { action(); });
        QObject::connect(b_print, &QPushButton::clicked, [this]() { print(); });
        QObject::connect(tabs, &QTabWidget::currentChanged, [this](int t) {
            if (t == t_find) { words(); fw->setFocus(); }
            else if (t == t_idx) kw->setFocus();
            else if (toc) toc->setFocus();
            buttons();
        });
        resize(440, 420);
        buttons();
    }
    void fill(const SmNode *n, QTreeWidgetItem *parent)
    {
        for (const SmNode &k : n->kids) {
            auto *it = parent ? new QTreeWidgetItem(parent) : new QTreeWidgetItem(toc);
            it->setText(0, k.name);
            bool book = k.locals.isEmpty();
            it->setIcon(0, hlp_icon(book ? 0 : 2));
            it->setData(0, Qt::UserRole, book ? QString() : k.locals.first());
            fill(&k, it);
            if (!book && it->childCount()) it->setExpanded(true);   /* (a topic with entries under it: open) */
        }
    }
    void buttons()
    {
        int t = tabs->currentIndex();
        QTreeWidgetItem *it = toc ? toc->currentItem() : nullptr;
        if (t == t_toc && it && it->data(0, Qt::UserRole).toString().isEmpty())
            b_main->setText(it->isExpanded() ? greek_or("Cl&ose", "Κλ&είσιμο") : greek_or("&Open", "Ά&νοιγμα"));
        else b_main->setText(greek_or("&Display", "&Εμφάνιση"));
    }
    void action()
    {
        int t = tabs->currentIndex();
        if (t == t_toc) open_item(toc->currentItem());
        else if (t == t_idx) show_key(idx->currentRow());
        else if (ftops->currentItem()) go_topic(ftops->currentItem()->data(Qt::UserRole).toInt());
    }
    void open_item(QTreeWidgetItem *it)
    {
        if (!it) return;
        QString loc = it->data(0, Qt::UserRole).toString();
        if (loc.isEmpty()) { it->setExpanded(!it->isExpanded()); buttons(); return; }
        QString ctx = loc.section('>', 0, 0), file = ctx.section('@', 1);
        ctx = ctx.section('@', 0, 0).trimmed();
        HlpDoc *t = hlp_other(d, file);
        quint32 off = 0;
        int ti = t ? hlp_topic_hash(t, hlp_hash(ctx), &off) : -1;
        if (ti < 0) { hlp_missing_topic(this); return; }
        hide();
        hlp_go(t, ti, off);
    }
    void find_key(const QString &s)
    {
        for (int i = 0; i < idx->count(); i++)
            if (idx->item(i)->text().startsWith(s, Qt::CaseInsensitive)) {
                idx->setCurrentRow(i);
                idx->scrollToItem(idx->item(i), QAbstractItemView::PositionAtTop);
                return;
            }
    }
    void show_key(int row)
    {
        if (row < 0 || row >= (int)d->kw.size()) return;
        std::vector<int> ts;
        std::vector<quint32> offs;
        for (quint32 o : d->kw[row].second) {
            int ti = hlp_topic_at(d, o);
            if (ti >= 0 && std::find(ts.begin(), ts.end(), ti) == ts.end()) { ts.push_back(ti); offs.push_back(o); }
        }
        if (ts.empty()) return;
        int k = 0;
        if (ts.size() > 1 && (k = hlp_pick(this, d, ts)) < 0) return;
        hide();
        hlp_go(d, ts[k], offs[k]);
    }
    static int hlp_pick(QWidget *parent, HlpDoc *d, const std::vector<int> &ts)   /* "Topics Found" */
    {
        QStringList names;
        for (int t : ts) names << hlp_topic_name(d, t);
        bool ok = false;
        QString pick = QInputDialog::getItem(parent, greek_or("Topics Found", "Θέματα που βρέθηκαν"),
                                             greek_or("Click a topic, then click Display.", "Επιλέξτε ένα θέμα και πατήστε Εμφάνιση."), names, 0, false, &ok);
        return ok ? names.indexOf(pick) : -1;
    }
    void words()                                /* Find: the words of every topic, once */
    {
        if (d->words_read) { if (!fwords->count()) find_words(); return; }
        d->words_read = true;
        QApplication::setOverrideCursor(Qt::WaitCursor);
        static const QRegularExpression sep("[^\\w']+");
        for (int ti = 0; ti < (int)d->tops.size(); ti++) {
            QString txt = d->tops[ti].title + " " + hlp_plain(d, ti);
            QSet<QString> seen;
            for (const QString &w : txt.toLower().split(sep, Qt::SkipEmptyParts)) {
                if (seen.contains(w)) continue;
                seen.insert(w);
                d->words[w].push_back(ti);
            }
        }
        QApplication::restoreOverrideCursor();
        find_words();
    }
    void find_words()
    {
        fwords->clear();
        QStringList typed = fw->text().toLower().split(' ', Qt::SkipEmptyParts);
        QString last = typed.isEmpty() ? QString() : typed.last();
        for (const auto &w : d->words) if (last.isEmpty() || w.first.startsWith(last)) fwords->addItem(w.first);
        find_topics();
    }
    void find_topics()
    {
        ftops->clear();
        QStringList typed = fw->text().toLower().split(' ', Qt::SkipEmptyParts);
        std::set<int> res;
        bool first = true;
        auto add = [&](const std::set<int> &s) {
            if (first) { res = s; first = false; return; }
            std::set<int> r;
            for (int t : s) if (res.count(t)) r.insert(t);
            res = r;
        };
        QList<QListWidgetItem *> sel = fwords->selectedItems();
        if (!sel.isEmpty()) {
            std::set<int> s;
            for (auto *it : sel) for (int t : d->words[it->text()]) s.insert(t);
            add(s);
        }
        for (int i = 0; i < typed.size() - (sel.isEmpty() ? 0 : 1); i++) {
            std::set<int> s;
            for (const auto &w : d->words) if (w.first.startsWith(typed[i])) s.insert(w.second.begin(), w.second.end());
            add(s);
        }
        if (first) res.clear();
        for (int t : res) {
            auto *it = new QListWidgetItem(hlp_topic_name(d, t), ftops);
            it->setData(Qt::UserRole, t);
        }
        fcount->setText(greek_or("%1 Topics Found", "Βρέθηκαν %1 θέματα").arg(ftops->count()));
    }
    void go_topic(int ti) { if (ti < 0) return; hide(); hlp_go(d, ti, d->tops[ti].off); }
    void print()                                /* the topic, or the topics of a book: a PDF file */
    {
        std::vector<int> ts;
        int t = tabs->currentIndex();
        std::function<void(QTreeWidgetItem *)> walk = [&](QTreeWidgetItem *it) {
            QString loc = it->data(0, Qt::UserRole).toString();
            if (!loc.isEmpty()) {
                HlpDoc *x = hlp_other(d, loc.section('>', 0, 0).section('@', 1));
                int ti = x == d ? hlp_topic_hash(d, hlp_hash(loc.section('>', 0, 0).section('@', 0, 0).trimmed())) : -1;
                if (ti >= 0) ts.push_back(ti);
            }
            for (int i = 0; i < it->childCount(); i++) walk(it->child(i));
        };
        if (t == t_toc && toc->currentItem()) walk(toc->currentItem());
        else if (t == t_idx && idx->currentRow() >= 0) { for (quint32 o : d->kw[idx->currentRow()].second) ts.push_back(hlp_topic_at(d, o)); }
        else if (t == t_find && ftops->currentItem()) ts.push_back(ftops->currentItem()->data(Qt::UserRole).toInt());
        ts.erase(std::remove(ts.begin(), ts.end(), -1), ts.end());
        if (ts.empty()) return;
        QString f = QFileDialog::getSaveFileName(this, greek_or("Print", "Εκτύπωση"),
            QFileInfo(d->file).absolutePath() + "/" + QFileInfo(d->file).completeBaseName() + ".pdf", "PDF (*.pdf)");
        if (f.isEmpty()) return;
        QTextDocument doc;
        QTextCursor c(&doc);
        for (size_t i = 0; i < ts.size(); i++) {
            QTextDocument one;
            hlp_render(d, ts[i], &one, nullptr, nullptr);
            if (i) { QTextBlockFormat pb; pb.setPageBreakPolicy(QTextFormat::PageBreak_AlwaysBefore); c.insertBlock(pb); }
            c.insertFragment(QTextDocumentFragment(&one));
            for (auto &p : d->pics) doc.addResource(QTextDocument::ImageResource, QUrl(QString("hlpbm:%1").arg(p.first)), p.second);
        }
        QPdfWriter pdf(f);
        pdf.setPageSize(QPageSize(QPageSize::A4));
        doc.print(&pdf);
    }
};
static std::map<HlpDoc *, QPointer<HlpTopics>> hlp_dlgs;
static void hlp_topics_dlg(HlpDoc *d, int tab, const QString &key)   /* tab: 0 Contents (else Index), 1 Index, 2 Find */
{
    QPointer<HlpTopics> &w = hlp_dlgs[d];
    if (!w) w = new HlpTopics(d);
    int t = tab == 2 ? w->t_find : tab == 1 || w->t_toc < 0 ? w->t_idx : w->t_toc;
    w->tabs->setCurrentIndex(t);
    if (t == w->t_idx) { w->kw->setText(key); w->find_key(key); w->kw->setFocus(); }
    help_raise(w);
}

/* ---------------------------------------------------------------- the main window of a help file */
class HlpWin : public QMainWindow {
public:
    HlpDoc *d;
    HlpBrowser *nsr, *body;
    QFrame *line;
    int cur = -1;
    quint32 curoff = 0;
    std::vector<std::pair<int, quint32>> back;
    QPushButton *b_back, *b_prev = nullptr, *b_next = nullptr;
    QMenu *m_book, *m_opt;
    QAction *a_top_def, *a_top_on, *a_top_off, *a_sys, *a_fs, *a_fn, *a_fl;
    explicit HlpWin(HlpDoc *doc) : QMainWindow(nullptr), d(doc)
    {
        setAttribute(Qt::WA_QuitOnClose, false);
        setWindowTitle(!d->caption.isEmpty() ? d->caption : !d->title.isEmpty() ? d->title : QString("Windows Help"));
        setWindowIcon(style()->standardIcon(QStyle::SP_DialogHelpButton));
        auto *mf = menuBar()->addMenu(greek_or("&File", "&Αρχείο"));
        QObject::connect(mf->addAction(greek_or("&Open...", "Ά&νοιγμα...")), &QAction::triggered, [this]() {
            QString f = QFileDialog::getOpenFileName(this, greek_or("Open", "Άνοιγμα"), QFileInfo(d->file).absolutePath(),
                                                     greek_or("Help Files (*.hlp *.HLP)", "Αρχεία βοήθειας (*.hlp *.HLP)"));
            if (f.isEmpty()) return;
            HlpDoc *o = hlp_doc(QFileInfo(f).absoluteFilePath());
            if (!o) { not_help(this, f); return; }
            hlp_go(o, hlp_contents(o), o->tops[hlp_contents(o)].off);
        });
        QObject::connect(mf->addAction(greek_or("&Print Topic...", "Ε&κτύπωση θέματος...")), &QAction::triggered, [this]() { print(); });
        mf->addSeparator();
        QObject::connect(mf->addAction(greek_or("E&xit", "Έ&ξοδος")), &QAction::triggered, [this]() { close(); });
        auto *me = menuBar()->addMenu(greek_or("&Edit", "&Επεξεργασία"));
        auto *a_copy = me->addAction(greek_or("&Copy", "Αντι&γραφή"));
        a_copy->setShortcut(QKeySequence::Copy);
        QObject::connect(a_copy, &QAction::triggered, [this]() { copy(); });
        me->addAction(greek_or("&Annotate...", "Σ&χολιασμός..."))->setEnabled(false);
        m_book = menuBar()->addMenu(greek_or("&Bookmark", "&Σελιδοδείκτης"));
        books();
        m_opt = menuBar()->addMenu(greek_or("&Options", "Επι&λογές"));
        options(m_opt);
        auto *mh = menuBar()->addMenu(greek_or("&Help", "&Βοήθεια"));
        QObject::connect(mh->addAction(greek_or("Help &Topics", "&Θέματα βοήθειας")), &QAction::triggered, [this]() { hlp_topics_dlg(d, 0, QString()); });
        QObject::connect(mh->addAction(greek_or("&Version...", "Έ&κδοση...")), &QAction::triggered, [this]() { version(this); });
        auto *tb = addToolBar("buttons");       /* the button bar of WinHelp */
        tb->setMovable(false);
        tb->setFloatable(false);
        auto *bw = new QWidget;
        auto *bl = new QHBoxLayout(bw);
        bl->setContentsMargins(2, 2, 2, 2);
        bl->setSpacing(2);
        auto button = [&](const QString &t) { auto *b = new QPushButton(t); b->setFocusPolicy(Qt::NoFocus); b->setMinimumWidth(70); bl->addWidget(b); return b; };
        QObject::connect(button(greek_or("Help &Topics", "&Θέματα βοήθειας")), &QPushButton::clicked, [this]() { hlp_topics_dlg(d, 0, QString()); });
        b_back = button(greek_or("&Back", "&Πίσω"));
        QObject::connect(b_back, &QPushButton::clicked, [this]() { go_back(); });
        auto *b_opt = button(greek_or("&Options", "Επι&λογές"));
        auto *pm = new QMenu(this);
        QObject::connect(pm->addAction(greek_or("&Annotate...", "Σ&χολιασμός...")), &QAction::triggered, []() {});
        pm->actions().last()->setEnabled(false);
        QObject::connect(pm->addAction(greek_or("&Copy", "Αντι&γραφή")), &QAction::triggered, [this]() { copy(); });
        QObject::connect(pm->addAction(greek_or("&Print Topic...", "Ε&κτύπωση θέματος...")), &QAction::triggered, [this]() { print(); });
        pm->addSeparator();
        options(pm);
        QObject::connect(b_opt, &QPushButton::clicked, [b_opt, pm]() { pm->exec(b_opt->mapToGlobal(QPoint(0, b_opt->height()))); });
        if (d->config.contains("BrowseButtons", Qt::CaseInsensitive)) {
            b_prev = button("&<<");
            b_next = button("&>>");
            QObject::connect(b_prev, &QPushButton::clicked, [this]() { browse(-1); });
            QObject::connect(b_next, &QPushButton::clicked, [this]() { browse(1); });
        }
        static const QRegularExpression cb("\\b(?:CreateButton|CB)\\s*\\(\\s*[`\"]([^'\"]*)['\"]\\s*,\\s*[`\"]([^'\"]*)['\"]\\s*,\\s*[`\"](.*?)['\"]\\s*\\)",
                                           QRegularExpression::CaseInsensitiveOption);
        auto cbi = cb.globalMatch(d->config);
        while (cbi.hasNext()) {
            auto m = cbi.next();
            QString mac = m.captured(3);
            QObject::connect(button(m.captured(2)), &QPushButton::clicked, [this, mac]() { hlp_run_macros(d, mac, this); });
        }
        bl->addStretch();
        tb->addWidget(bw);
        auto *cw = new QWidget;
        auto *v = new QVBoxLayout(cw);
        v->setContentsMargins(0, 0, 0, 0);
        v->setSpacing(0);
        nsr = new HlpBrowser(d, cw);
        nsr->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
        nsr->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
        line = new QFrame(cw);
        line->setFrameShape(QFrame::HLine);
        line->setFrameShadow(QFrame::Plain);
        body = new HlpBrowser(d, cw);
        v->addWidget(nsr);
        v->addWidget(line);
        v->addWidget(body, 1);
        setCentralWidget(cw);
        const HlpDoc::Win *mw = main_win();
        QRect a = QGuiApplication::primaryScreen() ? QGuiApplication::primaryScreen()->availableGeometry() : QRect(0, 0, 1024, 768);
        if (mw && (mw->flags & 0x78) == 0x78 && mw->w > 0 && mw->h > 0) {   /* the place of [WINDOWS] main: of 1024 x 1024 */
            setGeometry(a.left() + mw->x * a.width() / 1024, a.top() + mw->y * a.height() / 1024,
                        std::max(300, mw->w * a.width() / 1024), std::max(200, mw->h * a.height() / 1024));
        } else {
            resize(std::min(640, a.width() * 2 / 3), std::min(520, a.height() * 3 / 4));
            move(a.right() - width() - a.width() / 16, a.top() + a.height() / 8);
        }
    }
    const HlpDoc::Win *main_win() const
    {
        for (const auto &w : d->wins) if (w.name == "main") return &w;
        return nullptr;
    }
    static void not_help(QWidget *p, const QString &f)
    {
        QMessageBox::warning(p, "Windows Help", greek_or("The file %1 is not a Windows Help file, or the file is corrupted.",
                                                         "Το αρχείο %1 δεν είναι αρχείο Windows Help, ή είναι κατεστραμμένο.").arg(QDir::toNativeSeparators(f)));
    }
    static void version(QWidget *p)
    {
        QMessageBox::about(p, greek_or("Version", "Έκδοση"), greek_or("Windows Help for LinBASIC Studio\nA reader of the help files of WinHelp 3.1 and 4.",
                                                                     "Windows Help για το LinBASIC Studio\nΑναγνώστης των αρχείων βοήθειας του WinHelp 3.1 και 4."));
    }
    void options(QMenu *m)
    {
        auto *top = m->addMenu(greek_or("&Keep Help on Top", "&Βοήθεια πάντα στο προσκήνιο"));
        auto *g1 = new QActionGroup(m);
        auto add = [&](QMenu *mm, QActionGroup *g, const QString &t, bool on, std::function<void()> f) {
            QAction *a = mm->addAction(t); a->setCheckable(true); a->setChecked(on); g->addAction(a);
            QObject::connect(a, &QAction::triggered, f);
            return a;
        };
        add(top, g1, greek_or("&Default", "&Προεπιλογή"), true, [this]() { on_top(false); });
        add(top, g1, greek_or("&On Top", "&Στο προσκήνιο"), false, [this]() { on_top(true); });
        add(top, g1, greek_or("&Not On Top", "&Όχι στο προσκήνιο"), false, [this]() { on_top(false); });
        auto *fm = m->addMenu(greek_or("&Font", "&Γραμματοσειρά"));
        auto *g2 = new QActionGroup(m);
        add(fm, g2, greek_or("&Small", "&Μικρή"), hlp_scale < 1, [this]() { hlp_scale = 0.85; redraw(); });
        add(fm, g2, greek_or("&Normal", "&Κανονική"), hlp_scale == 1, [this]() { hlp_scale = 1.0; redraw(); });
        add(fm, g2, greek_or("&Large", "Μεγά&λη"), hlp_scale > 1, [this]() { hlp_scale = 1.25; redraw(); });
        QAction *sc = m->addAction(greek_or("&Use System Colors", "Χρήση χρωμάτων του σ&υστήματος"));
        sc->setCheckable(true);
        sc->setChecked(hlp_syscolors);
        QObject::connect(sc, &QAction::triggered, [this](bool on) { hlp_syscolors = on; redraw(); });
    }
    void on_top(bool on)
    {
        bool vis = isVisible();
        setWindowFlag(Qt::WindowStaysOnTopHint, on);
        if (vis) show();
    }
    void books()
    {
        m_book->clear();
        QObject::connect(m_book->addAction(greek_or("&Define...", "&Ορισμός...")), &QAction::triggered, [this]() {
            if (cur < 0) return;
            bool ok = false;
            QString n = QInputDialog::getText(this, greek_or("Bookmark Define", "Ορισμός σελιδοδείκτη"), greek_or("Bookmark Name:", "Όνομα σελιδοδείκτη:"),
                                              QLineEdit::Normal, hlp_topic_name(d, cur), &ok);
            if (!ok || n.trimmed().isEmpty()) return;
            d->marks.emplace_back(n.trimmed(), cur);
            books();
        });
        if (!d->marks.empty()) m_book->addSeparator();
        for (size_t i = 0; i < d->marks.size() && i < 9; i++) {
            int t = d->marks[i].second;
            QObject::connect(m_book->addAction(QString("&%1 %2").arg(i + 1).arg(d->marks[i].first)), &QAction::triggered,
                             [this, t]() { go(t, d->tops[t].off, true); });
        }
    }
    void go(int ti, quint32 off, bool push)
    {
        if (ti < 0 || ti >= (int)d->tops.size()) return;
        if (push && cur >= 0 && (cur != ti || curoff != off)) back.emplace_back(cur, curoff);
        cur = ti; curoff = off;
        draw();
    }
    void redraw() { if (cur >= 0) draw(); }
    void draw()
    {
        body->document()->clear();
        nsr->document()->clear();
        body->out = HlpOut();
        hlp_render(d, cur, body->document(), nsr->document(), &body->out);
        nsr->out.macros = body->out.macros;
        const HlpDoc::Win *mw = main_win();
        QColor white = hlp_syscolors ? QApplication::palette().base().color() : QColor(Qt::white);
        body->colors(mw && (mw->flags & 0x100) && !hlp_syscolors ? mw->rgb : white);
        nsr->colors(mw && (mw->flags & 0x200) && !hlp_syscolors ? mw->nsr : white);
        bool has = !nsr->document()->isEmpty();
        nsr->setVisible(has);
        line->setVisible(has);
        fit();
        b_back->setEnabled(!back.empty());
        const HlpDoc::Topic &t = d->tops[cur];
        if (b_prev) b_prev->setEnabled(t.bck != -1);
        if (b_next) b_next->setEnabled(t.fwd != -1);
        int pos = 0;
        for (const auto &m : body->out.marks) if (m.first <= curoff) pos = m.second;
        QPointer<HlpBrowser> b = body;
        QTimer::singleShot(0, [b, pos]() {
            if (!b) return;
            QTextCursor c(b->document());
            c.setPosition(std::min(pos, b->document()->characterCount() - 1));
            b->setTextCursor(c);
            int y = pos ? (int)b->document()->documentLayout()->blockBoundingRect(c.block()).top() : 0;
            b->verticalScrollBar()->setValue(y);
        });
    }
    void fit()                                  /* the non-scrolling region: as high as its text */
    {
        if (nsr->isHidden()) return;
        int w = centralWidget() && centralWidget()->width() > 50 ? centralWidget()->width() : width();
        nsr->document()->setTextWidth(std::max(50, w));
        nsr->setFixedHeight((int)std::ceil(nsr->document()->size().height()) + 2);
    }
    void resizeEvent(QResizeEvent *e) override { QMainWindow::resizeEvent(e); fit(); }
    void showEvent(QShowEvent *e) override { QMainWindow::showEvent(e); fit(); }
    void go_back()
    {
        if (back.empty()) return;
        auto b = back.back();
        back.pop_back();
        go(b.first, b.second, false);
    }
    void browse(int dir)
    {
        if (cur < 0) return;
        qint32 o = dir < 0 ? d->tops[cur].bck : d->tops[cur].fwd;
        if (o == -1) return;
        int ti = hlp_topic_at(d, (quint32)o);
        if (ti >= 0) go(ti, d->tops[ti].off, true);
    }
    void copy()
    {
        QString s = body->textCursor().hasSelection() ? body->textCursor().selectedText().replace(QChar(0x2029), '\n') : hlp_plain(d, cur);
        QApplication::clipboard()->setText(s);
    }
    void print()
    {
        if (cur < 0) return;
        QString f = QFileDialog::getSaveFileName(this, greek_or("Print Topic", "Εκτύπωση θέματος"),
            QFileInfo(d->file).absolutePath() + "/" + QFileInfo(d->file).completeBaseName() + ".pdf", "PDF (*.pdf)");
        if (f.isEmpty()) return;
        QTextDocument doc;
        hlp_render(d, cur, &doc, nullptr, nullptr);
        QPdfWriter pdf(f);
        pdf.setPageSize(QPageSize(QPageSize::A4));
        doc.print(&pdf);
    }
    void closeEvent(QCloseEvent *e) override
    {
        back.clear();
        cur = -1;
        if (hlp_dlgs[d]) hlp_dlgs[d]->close();
        QMainWindow::closeEvent(e);
    }
};
static std::map<HlpDoc *, QPointer<HlpWin>> hlp_wins;
static HlpWin *hlp_win(HlpDoc *d)
{
    QPointer<HlpWin> &w = hlp_wins[d];
    if (!w) w = new HlpWin(d);
    return w;
}
static void hlp_go(HlpDoc *d, int ti, quint32 off)
{
    HlpWin *w = hlp_win(d);
    w->go(ti, off, true);
    help_raise(w);
}
static void hlp_close_all()
{
    for (auto &p : hlp_dlgs) if (p.second) p.second->close();
    for (auto &p : hlp_wins) if (p.second) p.second->close();
}

/* ---------------------------------------------------------------- the macros of the help files */
static QStringList hlp_macro_split(const QString &s, QChar sep)   /* at the separators outside quotes and () */
{
    QStringList out;
    QString cur;
    int depth = 0;
    QChar q;
    for (int i = 0; i < s.size(); i++) {
        QChar c = s[i];
        if (!q.isNull()) { cur += c; if ((q == '`' && c == '\'') || (q == '"' && c == '"')) q = QChar(); continue; }
        if (c == '`' || c == '"') { q = c; cur += c; continue; }
        if (c == '(') depth++;
        if (c == ')') depth--;
        if (depth <= 0 && (c == sep || (sep == ';' && c == ':'))) { out << cur.trimmed(); cur.clear(); continue; }
        cur += c;
    }
    if (!cur.trimmed().isEmpty()) out << cur.trimmed();
    return out;
}
static QString hlp_unquote(QString a)
{
    a = a.trimmed();
    if (a.size() >= 2 && ((a.startsWith('`') && a.endsWith('\'')) || (a.startsWith('"') && a.endsWith('"')))) a = a.mid(1, a.size() - 2);
    return a;
}
static void hlp_run_macros(HlpDoc *d, const QString &macros, QWidget *from)
{
    for (const QString &m : hlp_macro_split(macros, ';')) {
        int p = m.indexOf('(');
        QString name = (p < 0 ? m : m.left(p)).trimmed().toLower();
        QStringList a;
        if (p >= 0) {
            QString in = m.mid(p + 1);
            int e = in.lastIndexOf(')');
            if (e >= 0) in = in.left(e);
            for (const QString &x : hlp_macro_split(in, ',')) a << hlp_unquote(x);
        }
        auto arg = [&](int k) { return k < a.size() ? a[k] : QString(); };
        auto jump = [&](HlpDoc *t, int ti, quint32 off, bool popup) {
            if (!t || ti < 0) { hlp_missing_topic(from); return; }
            if (popup) hlp_show_popup(t, ti, QCursor::pos()); else hlp_go(t, ti, off);
        };
        HlpWin *w = hlp_wins[d];
        if (name == "jumpid" || name == "ji" || name == "popupid" || name == "pi") {
            HlpDoc *t = hlp_other(d, arg(0));
            quint32 off = 0;
            int ti = t ? hlp_topic_hash(t, hlp_hash(arg(1)), &off) : -1;
            jump(t, ti, off, name[0] == 'p');
        } else if (name == "jumpcontext" || name == "jc" || name == "popupcontext" || name == "pc") {
            HlpDoc *t = hlp_other(d, arg(0));
            quint32 off = 0;
            int ti = t ? hlp_topic_id(t, arg(1).toUInt(), &off) : -1;
            jump(t, ti, off, name[0] == 'p');
        } else if (name == "jumpcontents" || name == "contents") {
            HlpDoc *t = name == "contents" ? d : hlp_other(d, arg(0));
            if (t) jump(t, hlp_contents(t), t->tops[hlp_contents(t)].off, false);
        } else if (name == "jumpkeyword" || name == "jk") {
            HlpDoc *t = hlp_other(d, arg(0));
            if (!t) continue;
            hlp_keywords(t);
            for (const auto &k : t->kw)
                if (k.first.compare(arg(1), Qt::CaseInsensitive) == 0 && !k.second.empty()) { jump(t, hlp_topic_at(t, k.second[0]), k.second[0], false); break; }
        } else if (name == "execfile" || name == "ef" || name == "execprogram" || name == "ep") {
            QString prog = arg(0), args = name.startsWith("execf") || name == "ef" ? arg(1) : QString();
            QUrl u = QUrl::fromUserInput(prog + (args.isEmpty() ? QString() : " " + args), QFileInfo(d->file).absolutePath());
            if (!QDesktopServices::openUrl(u))
                QMessageBox::warning(from, "Windows Help", greek_or("Unable to run the program %1.", "Δεν είναι δυνατή η εκτέλεση του %1.").arg(prog));
        } else if (name == "finder" || name == "fd") hlp_topics_dlg(d, 0, QString());
        else if (name == "search") hlp_topics_dlg(d, 1, QString());
        else if (name == "back") { if (w) w->go_back(); }
        else if (name == "next") { if (w) w->browse(1); }
        else if (name == "prev") { if (w) w->browse(-1); }
        else if (name == "print") { if (w) w->print(); }
        else if (name == "copytopic" || name == "ct") { if (w) w->copy(); }
        else if (name == "about") HlpWin::version(from);
        else if (name == "exit" || name == "closewindow" || name == "cw") { if (w) w->close(); }
    }
}

/* ---------------------------------------------------------------- the commands of the programs */
static void hlp_log_check(HlpDoc *d)            /* VBRT_HLP_CHECK: what the reader found (the corpus of the tests) */
{
    static std::set<HlpDoc *> done;
    if (!getenv("VBRT_HLP_CHECK") || done.count(d)) return;
    done.insert(d);
    hlp_keywords(d);
    hlp_cnt(d);
    int unk0 = d->unknown, chars = 0, pics = 0, badpics = 0;
    for (int i = 0; i < (int)d->tops.size(); i++) chars += hlp_plain(d, i).size();
    for (const auto &f : d->F) {
        if (!f.first.startsWith("|bm") && !f.first.startsWith("bm")) continue;
        pics++;
        if (hlp_picture(d, f.first.mid(f.first.startsWith('|') ? 3 : 2).toInt()).isNull()) badpics++;
    }
    int ctxok = 0;
    for (const auto &c : d->ctx) if (hlp_topic_at(d, c.second) >= 0) ctxok++;
    help_log(QString("check %1 minor %2 topics %3 contexts %4/%5 keywords %6 contents %7 pictures %8 bad %9 unknown %10 chars %11")
             .arg(QFileInfo(d->file).fileName()).arg(d->minor).arg(d->tops.size()).arg(ctxok).arg(d->ctx.size()).arg(d->kw.size())
             .arg(d->toc.kids.size()).arg(pics).arg(badpics).arg(d->unknown - unk0).arg(chars));
}
/* cmd, as WinHelp: 1 HELP_CONTEXT (0: the Contents topic), 3 HELP_CONTENTS, 0x0B HELP_FINDER, 0x101 HELP_KEY,
 * 0x105 HELP_PARTIALKEY */
static void hlp_open(const QString &full, int cmd, quint32 id, const QString &key, const char *why)
{
    QString name = QFileInfo(full).fileName();
    bool exists = QFileInfo(full).exists();
    HlpDoc *d = exists ? hlp_doc(full) : nullptr;
    int ti = -1;
    quint32 off = 0;
    std::vector<int> ts;
    std::vector<quint32> tos;
    if (d) {
        hlp_log_check(d);
        if (cmd == 1 && id) ti = hlp_topic_id(d, id, &off);
        else if (cmd == 1 || cmd == 3) ti = hlp_contents(d);
        else if (cmd == 0x101) {
            hlp_keywords(d);
            for (const auto &k : d->kw)
                if (k.first.compare(key, Qt::CaseInsensitive) == 0)
                    for (quint32 o : k.second) { int t = hlp_topic_at(d, o); if (t >= 0 && std::find(ts.begin(), ts.end(), t) == ts.end()) { ts.push_back(t); tos.push_back(o); } }
            if (ts.size() == 1) { ti = ts[0]; off = tos[0]; }
        }
        if (ti >= 0 && !off) off = d->tops[ti].off;
    }
    QString shown = !d ? QString("-") : cmd == 0x0B ? QString("[Help Topics]") : cmd == 0x105 || (cmd == 0x101 && ts.empty()) ? QString("[Index]")
                  : cmd == 0x101 && ts.size() > 1 ? QString("[Topics Found %1]").arg(ts.size()) : hlp_topic_name(d, ti);
    help_log(QString("%1 %2 %3%4").arg(why, name, cmd == 0x101 || cmd == 0x105 ? key + " " : QString(), shown));
    if (help_quiet()) { if (monkey_on) monkey_log(QString("Help %1: %2 %3").arg(why, name, shown)); return; }
    if (!exists) {                              /* WinHelp: the file is not there */
        auto r = QMessageBox::question(dialog_parent(), "Windows Help",
            greek_or("Cannot find the %1 file. Do you want to try to find this file yourself?",
                     "Δεν είναι δυνατή η εύρεση του αρχείου %1. Θέλετε να το αναζητήσετε εσείς;").arg(name), QMessageBox::Yes | QMessageBox::No);
        if (r != QMessageBox::Yes) return;
        QString f = QFileDialog::getOpenFileName(dialog_parent(), greek_or("Open", "Άνοιγμα"), QFileInfo(full).absolutePath(),
                                                 greek_or("Help Files (*.hlp *.HLP)", "Αρχεία βοήθειας (*.hlp *.HLP)"));
        if (!f.isEmpty()) hlp_open(QFileInfo(f).absoluteFilePath(), cmd, id, key, why);
        return;
    }
    if (!d) { HlpWin::not_help(dialog_parent(), full); return; }
    switch (cmd) {
    case 0x0B: hlp_topics_dlg(d, 0, QString()); break;
    case 0x105: hlp_topics_dlg(d, 1, key); break;
    case 0x101:
        if (ts.empty()) hlp_topics_dlg(d, 1, key);
        else if (ts.size() == 1) hlp_go(d, ti, off);
        else { int k = HlpTopics::hlp_pick(dialog_parent(), d, ts); if (k >= 0) hlp_go(d, ts[k], tos[k]); }
        break;
    default:
        if (ti < 0) { hlp_missing_topic(dialog_parent()); break; }
        hlp_go(d, ti, off);
    }
}
/* the pop-up of a context number (What's This, HELP_CONTEXTPOPUP); false: no such topic */
static bool hlp_popup(const QString &full, quint32 id, QPoint pt)
{
    HlpDoc *d = hlp_doc(full);
    int ti = d && id ? hlp_topic_id(d, id) : -1;
    if (ti < 0) return false;
    QString text = hlp_plain(d, ti).simplified();
    help_log(QString("popup %1 %2").arg(id).arg(text));
    if (help_quiet()) { if (monkey_on) monkey_log("What's This: " + text.left(80)); return true; }
    hlp_show_popup(d, ti, pt);
    return true;
}
