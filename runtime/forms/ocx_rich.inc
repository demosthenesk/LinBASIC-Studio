/* ocx_rich.inc - RichTextBox (RICHTX32.OCX): a QTextEdit that reads and writes the RTF of RichEdit.
 * Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
extern "C" {
extern int vb_codepage;
extern int vb_lcid;
vb_wchar vb_ansi2w(unsigned char c);
int vb_w2ansi(vb_wchar w);
}

/* the font name of a format: Qt 5 builds font() from FontFamily, Qt 6 from FontFamilies */
static void rt_family(QTextCharFormat &f, const QString &name)
{
#if QT_VERSION < QT_VERSION_CHECK(6, 0, 0)
    f.setFontFamily(name);
#endif
    f.setFontFamilies(QStringList{name});
}

/* properties of the text that Qt does not have: kept as user properties of the formats (twips) */
enum {
    RT_PROTECTED = QTextFormat::UserProperty + 1,      /* SelProtected */
    RT_OFFSET = QTextFormat::UserProperty + 2,         /* SelCharOffset, twips */
    RT_BULLET = QTextFormat::UserProperty + 3,         /* SelBullet (a paragraph) */
    RT_CHARSET = QTextFormat::UserProperty + 4,        /* \fcharset of the font in the RTF it came from (-1: none) */
};
static const double RT_TPP = 15.0;                     /* twips per pixel of the document */

static int rtf_codepage(int charset)
{
    switch (charset) {
    case 0: return 1252; case 161: return 1253; case 162: return 1254; case 204: return 1251;
    case 238: return 1250; case 186: return 1257; case 177: return 1255; case 178: return 1256;
    default: return 0;
    }
}
static int rtf_charset(int cp)
{
    switch (cp) {
    case 1253: return 161; case 1254: return 162; case 1251: return 204; case 1250: return 238;
    case 1257: return 186; case 1255: return 177; case 1256: return 178; default: return 0;
    }
}
/* a byte of a code page: the two that the runtime knows (1252, 1253), else the one of the system */
static QChar rtf_byte(unsigned char b, int cp)
{
    if (b < 0x80) return QChar(b);
    int keep = vb_codepage;
    if (cp == 1252 || cp == 1253) vb_codepage = cp;
    QChar r((ushort)vb_ansi2w(b));
    vb_codepage = keep;
    return r;
}

/* ---------------------------------------------------------------- RTF -> QTextDocument */
struct RtfReader {
    struct Font { QString name; int charset = -1; };
    struct St {
        int font = -1, fs = 24, cf = 0, vpos = 0, offset = 0, uc = 1;
        bool b = false, i = false, ul = false, strike = false, prot = false;
        int dest = 0;           /* 0 text, 1 fonttbl, 2 colortbl, 3 skipped, 4 pn (bullets), 5 field instruction */
    };
    struct Para { int align = 0, li = 0, ri = 0, fi = 0; bool bullet = false; QList<int> tabs; };
    QTextCursor cur;
    QMap<int, Font> fonts;
    QList<QColor> colors;       /* index 0: auto (invalid) */
    int deff = 0, ansicp = 0, curfont = -1, skip_uc = 0;
    bool color_set = false, pending_par = false, any = false;
    int cr = 0, cg = 0, cb = 0;
    St st;
    Para para;
    QList<St> stack;
    QString buf;

    int codepage() const
    {
        auto it = fonts.find(st.font >= 0 ? st.font : deff);
        int cp = it != fonts.end() && it->charset >= 0 ? rtf_codepage(it->charset) : 0;
        return cp ? cp : ansicp ? ansicp : vb_codepage;
    }
    QTextCharFormat charfmt() const
    {
        QTextCharFormat f;
        auto it = fonts.find(st.font >= 0 ? st.font : deff);
        if (it != fonts.end() && !it->name.isEmpty()) { rt_family(f, it->name); f.setProperty(RT_CHARSET, it->charset); }
        f.setFontPointSize(st.fs / 2.0);
        f.setFontWeight(st.b ? QFont::Bold : QFont::Normal);
        f.setFontItalic(st.i);
        f.setFontUnderline(st.ul);
        f.setFontStrikeOut(st.strike);
        if (st.cf > 0 && st.cf < colors.size() && colors[st.cf].isValid()) f.setForeground(colors[st.cf]);
        int off = st.offset;
        if (st.vpos == 1 && !off) off = 1; else if (st.vpos == 2 && !off) off = -1;
        if (off) {
            f.setVerticalAlignment(off > 0 ? QTextCharFormat::AlignSuperScript : QTextCharFormat::AlignSubScript);
            f.setProperty(RT_OFFSET, st.offset);
        }
        if (st.prot) f.setProperty(RT_PROTECTED, true);
        return f;
    }
    QTextBlockFormat blockfmt() const
    {
        QTextBlockFormat f;
        f.setAlignment(para.align == 1 ? Qt::AlignRight : para.align == 2 ? Qt::AlignHCenter :
                       para.align == 3 ? Qt::AlignJustify : Qt::AlignLeft);
        f.setLeftMargin(para.li / RT_TPP);
        f.setRightMargin(para.ri / RT_TPP);
        f.setTextIndent(para.fi / RT_TPP);
        if (!para.tabs.isEmpty()) {
            QList<QTextOption::Tab> t;
            for (int x : para.tabs) t.append(QTextOption::Tab(x / RT_TPP, QTextOption::LeftTab));
            f.setTabPositions(t);
        }
        if (para.bullet) f.setProperty(RT_BULLET, true);
        return f;
    }
    void flush()
    {
        if (buf.isEmpty()) return;
        if (pending_par) { cur.insertBlock(blockfmt(), charfmt()); pending_par = false; }
        else if (!any) cur.setBlockFormat(blockfmt());
        cur.insertText(buf, charfmt());
        any = true;
        buf.clear();
    }
    void text(QChar ch)
    {
        if (skip_uc > 0) { skip_uc--; return; }
        if (st.dest == 0) { buf += ch; return; }
        if (st.dest == 1) {
            if (ch == ';') { curfont = -1; return; }
            if (curfont >= 0) fonts[curfont].name += ch;
            return;
        }
        if (st.dest == 2 && ch == ';') {
            colors.append(color_set ? QColor(cr, cg, cb) : QColor());
            color_set = false; cr = cg = cb = 0;
        }
    }
    void par()
    {
        flush();
        if (pending_par) cur.insertBlock(blockfmt(), charfmt());
        else cur.setBlockFormat(blockfmt());
        pending_par = true;
        any = true;
    }
    void word(const QString &w, bool has, int n)
    {
        if (st.dest == 3) return;
        if (st.dest == 1) {                     /* the font table */
            if (w == "f") { curfont = n; fonts[n]; }
            else if (w == "fcharset" && curfont >= 0) fonts[curfont].charset = n;
            return;
        }
        if (st.dest == 2) {
            if (w == "red") { cr = n; color_set = true; } else if (w == "green") { cg = n; color_set = true; }
            else if (w == "blue") { cb = n; color_set = true; }
            return;
        }
        if (st.dest == 4) { if (w == "pnlvlblt") para.bullet = true; return; }
        static const QSet<QString> skipped = { "stylesheet", "info", "pict", "object", "header", "footer", "headerl",
            "headerr", "footerl", "footerr", "fldinst", "pntext", "pntxtb", "pntxta", "listtable", "listoverridetable",
            "rsidtbl", "themedata", "colorschememapping", "datastore", "latentstyles", "generator", "xmlnstbl",
            "mmathPr", "revtbl", "filetbl", "bkmkstart", "bkmkend", "footnote", "annotation", "comment" };
        if (skipped.contains(w)) { flush(); st.dest = 3; return; }
        if (w == "fonttbl") { flush(); st.dest = 1; return; }
        if (w == "colortbl") { flush(); st.dest = 2; colors.clear(); return; }
        if (w == "pn") { flush(); st.dest = 4; para.bullet = true; return; }
        if (w == "ansicpg") { ansicp = n; return; }
        if (w == "deff") { deff = n; return; }
        if (w == "uc") { st.uc = n; return; }
        if (w == "u") {
            if (st.dest == 0) buf += QChar((ushort)(n < 0 ? n + 65536 : n));
            skip_uc = st.uc;
            return;
        }
        if (w == "par") { par(); return; }
        if (w == "line") { buf += QChar::LineSeparator; return; }
        if (w == "tab") { buf += '\t'; return; }
        if (w == "emdash") { buf += QChar(0x2014); return; } if (w == "endash") { buf += QChar(0x2013); return; }
        if (w == "bullet") { buf += QChar(0x2022); return; }
        if (w == "lquote") { buf += QChar(0x2018); return; } if (w == "rquote") { buf += QChar(0x2019); return; }
        if (w == "ldblquote") { buf += QChar(0x201C); return; } if (w == "rdblquote") { buf += QChar(0x201D); return; }
        /* formatting: the text so far keeps the old one */
        flush();
        bool on = !has || n != 0;
        if (w == "plain") { st.b = st.i = st.ul = st.strike = st.prot = false; st.fs = 24; st.cf = 0; st.vpos = 0; st.offset = 0; st.font = deff; }
        else if (w == "f") st.font = n;
        else if (w == "fs") st.fs = n;
        else if (w == "b") st.b = on;
        else if (w == "i") st.i = on;
        else if (w == "ul") st.ul = on;
        else if (w == "ulnone") st.ul = false;
        else if (w == "strike") st.strike = on;
        else if (w == "cf") st.cf = n;
        else if (w == "super") st.vpos = 1;
        else if (w == "sub") st.vpos = 2;
        else if (w == "nosupersub") { st.vpos = 0; st.offset = 0; }
        else if (w == "up") st.offset = n * 10;         /* half-points -> twips */
        else if (w == "dn") st.offset = -n * 10;
        else if (w == "protect") st.prot = on;
        else if (w == "pard") para = Para();
        else if (w == "ql") para.align = 0;
        else if (w == "qr") para.align = 1;
        else if (w == "qc") para.align = 2;
        else if (w == "qj") para.align = 3;
        else if (w == "li") para.li = n;
        else if (w == "ri") para.ri = n;
        else if (w == "fi") para.fi = n;
        else if (w == "tx") para.tabs.append(n);
    }
    void parse(const QByteArray &d)
    {
        int i = 0, len = d.size();
        bool star = false;
        while (i < len) {
            char ch = d[i];
            if (ch == '{') { flush(); stack.append(st); i++; continue; }
            if (ch == '}') {
                flush();
                if (!stack.isEmpty()) st = stack.takeLast();
                i++; continue;
            }
            if (ch == '\r' || ch == '\n') { i++; continue; }
            if (ch != '\\') { text(rtf_byte((unsigned char)ch, codepage())); i++; continue; }
            i++;
            if (i >= len) break;
            ch = d[i];
            if (ch == '\'') {
                bool ok = false;
                int b = d.mid(i + 1, 2).toInt(&ok, 16);
                i += 3;
                if (ok) text(rtf_byte((unsigned char)b, codepage()));
                continue;
            }
            if (ch == '*') { star = true; i++; continue; }
            if (!isalpha((unsigned char)ch)) {           /* \\ \{ \} \~ \- \_ */
                i++;
                if (ch == '~') text(QChar(0xA0));
                else if (ch == '_') text(QChar(0x2011));
                else if (ch == '-') {}
                else if (ch == '\r' || ch == '\n') { if (st.dest == 0) par(); }
                else text(QChar(ch));
                continue;
            }
            int s = i;
            while (i < len && isalpha((unsigned char)d[i])) i++;
            QString w = QString::fromLatin1(d.mid(s, i - s));
            bool has = false; int n = 0;
            if (i < len && (d[i] == '-' || isdigit((unsigned char)d[i]))) {
                int t = i; i++;
                while (i < len && isdigit((unsigned char)d[i])) i++;
                n = d.mid(t, i - t).toInt(); has = true;
            }
            if (i < len && d[i] == ' ') i++;
            if (star) {
                star = false;
                if (w != "pn") { flush(); st.dest = 3; continue; }
            }
            word(w, has, n);
        }
        flush();
    }
};
/* the document's paragraphs marked as bulleted get a list (the dot of the paragraph) */
static void rtf_bullets(QTextDocument *doc)
{
    QTextList *list = nullptr;
    for (QTextBlock b = doc->begin(); b.isValid(); b = b.next()) {
        bool on = b.blockFormat().boolProperty(RT_BULLET);
        if (on && !b.textList()) {
            if (!list) { QTextCursor c(b); QTextListFormat lf; lf.setStyle(QTextListFormat::ListDisc); list = c.createList(lf); }
            else list->add(b);
        } else if (!on) {
            if (b.textList()) b.textList()->remove(b);
            list = nullptr;
        }
    }
}
static bool rtf_read(QTextDocument *doc, const QByteArray &data, QTextCursor at)
{
    if (!data.trimmed().startsWith("{\\rtf")) return false;
    RtfReader r;
    r.cur = at;
    r.parse(data);
    rtf_bullets(doc);
    return true;
}

/* ---------------------------------------------------------------- QTextDocument -> RTF */
/* the font of a run; a size in pixels (the fonts of the forms) becomes points, 96 dpi */
static QFont rt_font(const QTextCharFormat &f, const QFont &base)
{
    QFont q = f.font().resolve(base);
    if (q.pointSizeF() <= 0 && q.pixelSize() > 0) q.setPointSizeF(q.pixelSize() * 0.75);
    return q;
}
static void rtf_text(QByteArray &o, const QString &t)
{
    for (QChar ch : t) {
        ushort u = ch.unicode();
        if (u == '\\' || u == '{' || u == '}') { o += '\\'; o += (char)u; }
        else if (u == '\t') o += "\\tab ";
        else if (u == QChar::LineSeparator) o += "\\line ";
        else if (u == 0xA0) o += "\\~";
        else if (u < 0x80) o += (char)u;
        else {
            int a = vb_w2ansi(u);
            if (a >= 0) o += QString::asprintf("\\'%02x", a).toLatin1();
            else o += QString::asprintf("\\u%d?", (short)u).toLatin1();
        }
    }
}
/* RTF of the characters [a, b) of the document, as RichEdit writes it */
static QByteArray rtf_write(QTextDocument *doc, int a, int b, bool whole)
{
    /* as RichEdit 1.0 (the RichTextBox of VB6) writes it:
     *   {\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset161 MS Sans Serif;}}CRLF
     *   \viewkind4\uc1\pard\lang1032\f0\fs17 line one CRLF\par line two CRLF\par }CRLF
     * the RTF of a selection has no \viewkind4 and no last \par */
    QFont base = doc->defaultFont();
    struct FontE { QString name; int charset; };
    QList<FontE> fonts;
    QList<QColor> colors;
    struct Run { QString text; QTextCharFormat f; };
    struct Par { QTextBlockFormat bf; bool bullet; QList<Run> runs; bool end; QTextCharFormat mark; };
    QList<Par> pars;
    for (QTextBlock bl = doc->findBlock(a); bl.isValid() && bl.position() <= b; bl = bl.next()) {
        Par p{bl.blockFormat(), bl.textList() != nullptr || bl.blockFormat().boolProperty(RT_BULLET), {}, false,
              bl.next().isValid() ? bl.next().charFormat() : bl.charFormat()};
        int bend = bl.position() + bl.length() - 1;            /* its paragraph mark */
        for (auto it = bl.begin(); !it.atEnd(); ++it) {
            QTextFragment fr = it.fragment();
            int s = std::max(a, fr.position()), e = std::min(b, fr.position() + fr.length());
            if (s >= e) continue;
            p.runs.append({fr.text().mid(s - fr.position(), e - s), fr.charFormat()});
        }
        if (p.runs.isEmpty()) p.runs.append({QString(), bl.charFormat()});
        p.end = bend < b;
        pars.append(p);
        if (bend >= b) break;
    }
    if (pars.isEmpty()) pars.append({QTextBlockFormat(), false, {{QString(), QTextCharFormat()}}, false, QTextCharFormat()});
    auto font_ix = [&](const QTextCharFormat &f) {
        QString n = rt_font(f, base).family();
        for (int k = 0; k < fonts.size(); k++) if (fonts[k].name == n) return k;
        /* the control's own font has the charset of the system; a font of an RTF its own, or none */
        int cs = f.hasProperty(RT_CHARSET) ? f.intProperty(RT_CHARSET) : rtf_charset(vb_codepage);
        fonts.append({n, cs});
        return (int)fonts.size() - 1;
    };
    bool bullets = false;
    for (auto &p : pars) bullets |= p.bullet;
    for (auto &p : pars) {
        QList<QTextCharFormat> fs;
        for (auto &r : p.runs) fs.append(r.f);
        if (p.end || whole) fs.append(p.mark);
        for (auto &cf : fs) {
            font_ix(cf);
            if (cf.hasProperty(QTextFormat::ForegroundBrush)) {
                QColor c = cf.foreground().color();
                if (!colors.contains(c)) colors.append(c);
            }
        }
    }
    int symbol = -1;                            /* the font of the bullets, last in the table */
    if (bullets) { fonts.append({"Symbol", 2}); symbol = (int)fonts.size() - 1; }
    QByteArray o = "{\\rtf1\\ansi\\deff0{\\fonttbl";
    for (int k = 0; k < fonts.size(); k++) {
        o += "{\\f" + QByteArray::number(k) + "\\fnil";
        if (fonts[k].charset >= 0) o += "\\fcharset" + QByteArray::number(fonts[k].charset);
        o += " ";
        rtf_text(o, fonts[k].name);
        o += ";}";
    }
    o += "}\r\n";
    if (!colors.isEmpty()) {
        o += "{\\colortbl ;";
        for (auto &c : colors) o += "\\red" + QByteArray::number(c.red()) + "\\green" + QByteArray::number(c.green()) + "\\blue" + QByteArray::number(c.blue()) + ";";
        o += "}\r\n";
    }
    if (whole) o += "\\viewkind4";
    o += "\\uc1";
    int f = -1, fs = -1, cf = 0, off = 0; bool bo = false, it = false, ul = false, sk = false, pr = false, lang = false;
    QByteArray lastpar;
    for (int pi = 0; pi < pars.size(); pi++) {
        Par &p = pars[pi];
        QByteArray pw = "\\pard";
        if (p.bullet) {
            QByteArray sf = QByteArray::number(symbol);
            pw += "{\\pntext\\f" + sf + "\\'B7\\tab}{\\*\\pn\\pnlvlblt\\pnf" + sf + "\\pnindent0{\\pntxtb\\'B7}}";
        }
        int li = (int)std::lround(p.bf.leftMargin() * RT_TPP), ri = (int)std::lround(p.bf.rightMargin() * RT_TPP), fi = (int)std::lround(p.bf.textIndent() * RT_TPP);
        if (fi) pw += "\\fi" + QByteArray::number(fi);
        if (li) pw += "\\li" + QByteArray::number(li);
        if (ri) pw += "\\ri" + QByteArray::number(ri);
        Qt::Alignment al = p.bf.alignment() & Qt::AlignHorizontal_Mask;
        if (al == Qt::AlignRight) pw += "\\qr"; else if (al == Qt::AlignHCenter) pw += "\\qc"; else if (al == Qt::AlignJustify) pw += "\\qj";
        for (auto &t : p.bf.tabPositions()) pw += "\\tx" + QByteArray::number((int)std::lround(t.position * RT_TPP));
        if (pi == 0 || pw != lastpar) { o += pw; lastpar = pw; }
        if (!lang) { o += "\\lang" + QByteArray::number(vb_lcid); lang = true; }
        auto fmt = [&](const QTextCharFormat &rf) {
            QFont q = rt_font(rf, base);
            QByteArray w;
            int nf = font_ix(rf);
            int nfs = (int)std::lround(q.pointSizeF() * 2);
            int ncf = rf.hasProperty(QTextFormat::ForegroundBrush) ? (int)colors.indexOf(rf.foreground().color()) + 1 : 0;
            int noff = rf.hasProperty(RT_OFFSET) ? rf.intProperty(RT_OFFSET) :
                       rf.verticalAlignment() == QTextCharFormat::AlignSuperScript ? 1 : rf.verticalAlignment() == QTextCharFormat::AlignSubScript ? -1 : 0;
            bool npr = rf.boolProperty(RT_PROTECTED);
            if (ncf != cf) { w += ncf ? "\\cf" + QByteArray::number(ncf) : QByteArray("\\cf0"); cf = ncf; }
            if (q.underline() != ul) { w += q.underline() ? "\\ul" : "\\ulnone"; ul = q.underline(); }
            if (q.bold() != bo) { w += q.bold() ? "\\b" : "\\b0"; bo = q.bold(); }
            if (q.italic() != it) { w += q.italic() ? "\\i" : "\\i0"; it = q.italic(); }
            if (q.strikeOut() != sk) { w += q.strikeOut() ? "\\strike" : "\\strike0"; sk = q.strikeOut(); }
            if (noff != off) {
                if (noff == 1) w += "\\super";
                else if (noff == -1) w += "\\sub";
                else if (noff == 0 && (off == 1 || off == -1)) w += "\\nosupersub";
                else w += (noff >= 0 ? "\\up" : "\\dn") + QByteArray::number(std::abs(noff) / 10);   /* \up0 ends \up4 */
                off = noff;
            }
            if (npr != pr) { w += npr ? "\\protect" : "\\protect0"; pr = npr; }
            if (nf != f) { w += "\\f" + QByteArray::number(nf); f = nf; }
            if (nfs != fs) { w += "\\fs" + QByteArray::number(nfs); fs = nfs; }
            if (!w.isEmpty()) { o += w; o += ' '; }
        };
        for (auto &r : p.runs) {
            fmt(r.f);
            rtf_text(o, r.text);
        }
        if (p.end || (whole && pi == pars.size() - 1)) {
            fmt(p.mark);                        /* the paragraph mark has a format of its own */
            o += "\r\n\\par ";
        }
    }
    o += "}\r\n";
    return o;
}

/* ---------------------------------------------------------------- the control */
class RichEdit : public QTextEdit {
public:
    explicit RichEdit(QWidget *parent) : QTextEdit(parent) {}
};
struct OcxRichText : Ocx {
    QPointer<QTextEdit> te;
    QString filename, last;
    int maxlength = 0, scrollbars = 0;
    bool multiline = true, hidesel = true, autoverb = false, loading = false, quiet = false;
    double rightmargin = 0;

    QWidget *make(QWidget *parent) override
    {
        te = new RichEdit(parent);
        te->setAcceptRichText(false);
        te->setTabChangesFocus(false);
        vbobj self = c->self;
        QObject::connect(te->document(), &QTextDocument::contentsChanged, [self]() {
            Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxRichText *>(x->ocx) : nullptr;
            if (!o || o->loading || o->quiet) return;
            QString t = o->te->toPlainText();
            if (t == o->last) return;           /* a change of the format only: no Change */
            o->last = t;
            fire_ctl(x, "change");
        });
        QObject::connect(te, &QTextEdit::cursorPositionChanged, [self]() { Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxRichText *>(x->ocx) : nullptr; if (o && !o->loading && !o->quiet) fire_ctl(x, "selchange"); });
        return te;
    }
    QTextDocument *doc() { return te->document(); }
    /* VB6 counts the end of a paragraph as two characters (CR LF), Qt as one */
    int vbpos(int qp) { QString q = te->toPlainText(); return qp + (int)q.left(qp).count('\n'); }
    int qpos(int vp)
    {
        QString q = te->toPlainText();
        int i = 0, v = 0;
        while (i < q.size() && v < vp) { v += q[i] == '\n' ? 2 : 1; i++; }
        return i;
    }
    QString vbtext(QString t) { return t.replace(QChar::ParagraphSeparator, '\n').replace('\n', QStringLiteral("\r\n")); }
    void quiet_set(const std::function<void()> &f)
    {
        quiet = true; f(); quiet = false;
        last = te->toPlainText();
    }
    void set_text(const QString &t)
    {
        QString q = QString(t).replace(QStringLiteral("\r\n"), QStringLiteral("\n")).replace('\r', '\n');
        if (q.isEmpty() && doc()->isEmpty()) return;      /* nothing replaced by nothing: no Change */
        QTextCharFormat cf = te->currentCharFormat();
        if (!doc()->isEmpty()) { QTextCursor first(doc()); first.setPosition(1); cf = first.charFormat(); }
        loading = true;
        te->clear();                            /* RichEdit: the new text takes the format of the first character */
        te->setCurrentCharFormat(cf);
        QTextCursor tc(doc());
        tc.insertText(q, cf);
        tc.select(QTextCursor::Document);
        tc.setBlockCharFormat(cf);              /* the paragraph marks too */
        tc.setPosition(0);
        te->setTextCursor(tc);
        loading = false;
        changed();
    }
    void changed()                              /* the program changed the text or its format: Change */
    {
        last = te->toPlainText();
        if (!quiet) fire_ctl(c, "change");
    }
    bool set_rtf(const QByteArray &data, bool whole)
    {
        loading = true;
        QTextCursor tc = te->textCursor();
        bool ok;
        if (whole) {
            if (!data.trimmed().startsWith("{\\rtf")) {         /* not RTF: the text, as Text = does */
                loading = false;
                QString t;
                for (char ch : data) t += rtf_byte((unsigned char)ch, vb_codepage);
                set_text(t);
                return false;
            }
            te->clear();
            QTextCursor at(doc());
            ok = rtf_read(doc(), data, at);
            at.setPosition(0);
            te->setTextCursor(at);
        } else {
            QTextDocument tmp;
            tmp.setDefaultFont(doc()->defaultFont());
            QTextCursor at(&tmp);
            ok = rtf_read(&tmp, data, at);
            if (!ok) at.insertText(QString::fromLatin1(data).replace(QStringLiteral("\r\n"), QStringLiteral("\n")));
            tc.insertFragment(QTextDocumentFragment(&tmp));
            te->setTextCursor(tc);
        }
        loading = false;
        changed();
        return ok;
    }
    void load(const char *const *p) override
    {
        QString v;
        multiline = propn(p, "MultiLine", -1) != 0;
        scrollbars = (int)propn(p, "ScrollBars", 0);
        maxlength = (int)propn(p, "MaxLength", 0);
        hidesel = propn(p, "HideSelection", -1) != 0;
        autoverb = propn(p, "AutoVerbMenu", 0) != 0;
        rightmargin = propn(p, "RightMargin", 0);
        te->setReadOnly(propn(p, "Locked", 0) != 0);
        te->setVerticalScrollBarPolicy(scrollbars == 2 || scrollbars == 3 ? Qt::ScrollBarAsNeeded : Qt::ScrollBarAlwaysOff);
        te->setHorizontalScrollBarPolicy(scrollbars == 1 || scrollbars == 3 ? Qt::ScrollBarAsNeeded : Qt::ScrollBarAlwaysOff);
        te->setLineWrapMode(multiline && !(scrollbars == 1 || scrollbars == 3) ? QTextEdit::WidgetWidth : QTextEdit::NoWrap);
        if (propn(p, "BorderStyle", 1) == 0) te->setFrameStyle(QFrame::NoFrame);
        quiet_set([&]() {
            if (prop(p, "TextRTF", &v)) {
                QByteArray d;
                for (QChar ch : v) { int a = vb_w2ansi(ch.unicode()); d += a >= 0 ? (char)a : '?'; }
                set_rtf(d, true);
            } else if (prop(p, "Text", &v)) set_text(v);
        });
        if (prop(p, "FileName", &v) && !v.isEmpty()) { filename = v; quiet_set([&]() { load_file(v, 0, false); }); }
    }
    bool load_file(const QString &path, int type, bool raise)
    {
        QFile f(native(path));
        if (!f.open(QIODevice::ReadOnly)) { if (raise) vb_raise(75); return false; }
        QByteArray d = f.readAll();
        if (type == 0 && !d.trimmed().startsWith("{\\rtf")) {
            changed();                          /* VB6 fires Change for the failed load too */
            if (raise) vb_raise_full(321, vs(QString("RichTextBox")), vs(QString("Invalid file format")));
            return false;
        }
        if (type == 0) set_rtf(d, true);
        else {
            QString s;
            for (char ch : d) s += rtf_byte((unsigned char)ch, vb_codepage);
            set_text(s);
        }
        return true;
    }
    QByteArray ansi(const QString &s)
    {
        QByteArray d;
        for (QChar ch : s) { int a = vb_w2ansi(ch.unicode()); d += a >= 0 ? (char)a : '?'; }
        return d;
    }
    /* the selection: Qt positions */
    void sel(int *a, int *b) { QTextCursor tc = te->textCursor(); *a = tc.selectionStart(); *b = tc.selectionEnd(); }
    /* a property of the characters of the selection: Null when they differ */
    vbvar sel_char(const std::function<QVariant(const QTextCharFormat &)> &get)
    {
        int a, b; sel(&a, &b);
        if (a == b) return variant(get(te->currentCharFormat()));
        QVariant first; bool have = false;
        for (QTextBlock bl = doc()->findBlock(a); bl.isValid() && bl.position() < b; bl = bl.next())
            for (auto it = bl.begin(); !it.atEnd(); ++it) {
                QTextFragment fr = it.fragment();
                if (fr.position() + fr.length() <= a || fr.position() >= b) continue;
                QVariant v = get(fr.charFormat());
                if (!have) { first = v; have = true; } else if (v != first) return vb_v_null();
            }
        return have ? variant(first) : variant(get(te->currentCharFormat()));
    }
    vbvar sel_block(const std::function<QVariant(const QTextBlock &)> &get)
    {
        int a, b; sel(&a, &b);
        QVariant first; bool have = false;
        for (QTextBlock bl = doc()->findBlock(a); bl.isValid() && (bl.position() <= b || !have); bl = bl.next()) {
            if (have && bl.position() == b && b > a) break;
            QVariant v = get(bl);
            if (!have) { first = v; have = true; } else if (v != first) return vb_v_null();
        }
        return variant(first);
    }
    vbvar variant(const QVariant &v)
    {
        switch (v.userType()) {
        case QMetaType::Bool: return vbool(v.toBool());
        case QMetaType::Int: return vb_v_l(v.toInt());
        case QMetaType::Double: return vb_v_f((float)v.toDouble());
        case QMetaType::QString: return vstr(v.toString());
        default: return vb_v_null();
        }
    }
    void merge_char(const QTextCharFormat &f)
    {
        QTextCursor tc = te->textCursor();
        loading = true;
        if (tc.hasSelection()) tc.mergeCharFormat(f);
        te->mergeCurrentCharFormat(f);
        loading = false;
        if (tc.hasSelection()) changed();       /* VB6: a format at the insertion point alone is no Change */
    }
    void each_block(const std::function<void(QTextCursor &, QTextBlockFormat &)> &f)
    {
        int a, b; sel(&a, &b);
        loading = true;
        for (QTextBlock bl = doc()->findBlock(a); bl.isValid(); bl = bl.next()) {
            if (bl.position() > b || (bl.position() == b && b > a)) break;
            QTextCursor tc(bl);
            QTextBlockFormat bf = bl.blockFormat();
            f(tc, bf);
            tc.setBlockFormat(bf);
        }
        rtf_bullets(doc());
        loading = false;
        changed();
    }
    QFont fnt(const QTextCharFormat &f) { return rt_font(f, doc()->defaultFont()); }
    /* the control gives the sizes of paragraphs in HIMETRIC (0.01 mm): 300 twips come back as 299.9055 */
    double tw(double px)
    {
        float hm = (float)std::round(px * RT_TPP * 2540 / 1440);
        return from_twips_ctl(c, (double)(hm * (float)(1440.0 / 2540.0)));    /* in Single, as VB6 */
    }
    double px(vbvar v) { return std::round(to_twips_ctl(c, vb_v2d(v))) / RT_TPP; }      /* whole twips */
    double px_pixels(vbvar v) { return std::round(to_twips_ctl(c, vb_v2d(v)) / RT_TPP); } /* whole pixels: the right indent, tabs */
    /* the characters of a set, for Span and UpTo */
    static QChar pch(const QString &q, int i) { QChar c = q.at(i); return c == QChar('\n') ? QChar(QChar::ParagraphSeparator) : c; }
    static bool in_set(QChar ch, const QString &set, bool negate) { bool in = set.contains(ch == QChar::ParagraphSeparator ? QChar('\r') : ch); return negate ? !in : in; }

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        if (!te) { *found = 0; return vb_v_empty(); }
        QTextCursor tc = te->textCursor();
        if (n == "text") { if (!set) return vstr(vbtext(te->toPlainText())); set_text(qs(vb_v2s(val))); return vb_v_empty(); }
        if (n == "textrtf") {
            if (!set) { QByteArray d = rtf_write(doc(), 0, doc()->characterCount() - 1, true); QString s; for (char ch : d) s += rtf_byte((unsigned char)ch, vb_codepage); return vstr(s); }
            set_rtf(ansi(qs(vb_v2s(val))), true);
            return vb_v_empty();
        }
        if (n == "selrtf") {
            int a, b; sel(&a, &b);
            if (!set) { QByteArray d = rtf_write(doc(), a, b, false); QString s; for (char ch : d) s += rtf_byte((unsigned char)ch, vb_codepage); return vstr(s); }
            QString v = qs(vb_v2s(val));
            if (v.isEmpty() && !te->textCursor().hasSelection()) return vb_v_empty();
            set_rtf(ansi(v), false);
            return vb_v_empty();
        }
        if (n == "selstart" || n == "sellength" || n == "seltext") {
            int a = tc.selectionStart(), b = tc.selectionEnd();
            if (!set) {
                if (n == "selstart") return vb_v_l(vbpos(a));
                if (n == "sellength") return vb_v_l(vbpos(b) - vbpos(a));
                return vstr(vbtext(tc.selectedText()));
            }
            if (n == "selstart") {
                int v = vb_v2l(val); if (vb_errpend) return vb_v_empty();
                if (v < 0) { vb_raise(380); return vb_v_empty(); }
                tc.setPosition(qpos(v)); te->setTextCursor(tc); return vb_v_empty();
            }
            if (n == "sellength") {
                int v = vb_v2l(val); if (vb_errpend) return vb_v_empty();
                if (v < 0) { vb_raise(380); return vb_v_empty(); }
                tc.setPosition(a); tc.setPosition(qpos(vbpos(a) + v), QTextCursor::KeepAnchor); te->setTextCursor(tc); return vb_v_empty();
            }
            QString ins = qs(vb_v2s(val)).replace(QStringLiteral("\r\n"), QStringLiteral("\n")).replace('\r', '\n');
            if (ins.isEmpty() && !tc.hasSelection()) return vb_v_empty();
            loading = true;
            tc.insertText(ins, te->currentCharFormat());
            te->setTextCursor(tc);
            loading = false;
            changed();
            return vb_v_empty();
        }
        /* the characters of the selection */
        auto setchar = [&](const std::function<void(QTextCharFormat &)> &f) { QTextCharFormat cf; f(cf); merge_char(cf); return vb_v_empty(); };
        if (n == "selbold") { if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(fnt(f).bold()); }); bool v = vb_v2bool(val) != 0; return setchar([&](QTextCharFormat &f) { f.setFontWeight(v ? QFont::Bold : QFont::Normal); }); }
        if (n == "selitalic") { if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(fnt(f).italic()); }); bool v = vb_v2bool(val) != 0; return setchar([&](QTextCharFormat &f) { f.setFontItalic(v); }); }
        if (n == "selunderline") { if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(fnt(f).underline()); }); bool v = vb_v2bool(val) != 0; return setchar([&](QTextCharFormat &f) { f.setFontUnderline(v); }); }
        if (n == "selstrikethru") { if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(fnt(f).strikeOut()); }); bool v = vb_v2bool(val) != 0; return setchar([&](QTextCharFormat &f) { f.setFontStrikeOut(v); }); }
        if (n == "selfontname") { if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(fnt(f).family()); }); QString v = qs(vb_v2s(val)); return setchar([&](QTextCharFormat &f) { rt_family(f, v); }); }
        if (n == "selfontsize") {
            if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(fnt(f).pointSizeF()); });
            double v = vb_v2d(val); if (vb_errpend) return vb_v_empty();
            if (v > 1638 || v < 0) { vb_raise(380); return vb_v_empty(); }
            if (v == 0) { if (te->textCursor().hasSelection()) changed(); return vb_v_empty(); }   /* no error, the size stays */
            return setchar([&](QTextCharFormat &f) { f.setFontPointSize(std::round(v * 20) / 20); });
        }
        if (n == "selcolor") {
            if (!set) return sel_char([&](const QTextCharFormat &f) {
                if (!f.hasProperty(QTextFormat::ForegroundBrush)) return QVariant(0);
                QColor q = f.foreground().color(); return QVariant(q.red() | q.green() << 8 | q.blue() << 16); });
            long v = vb_v2l(val);
            QColor q = qcolor(v);
            return setchar([&](QTextCharFormat &f) { f.setForeground(q); });
        }
        if (n == "selcharoffset") {
            if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(f.intProperty(RT_OFFSET)); });
            int v = (int)std::lround(to_twips_ctl(c, vb_v2d(val)));
            return setchar([&](QTextCharFormat &f) {
                f.setProperty(RT_OFFSET, v);
                f.setVerticalAlignment(v > 0 ? QTextCharFormat::AlignSuperScript : v < 0 ? QTextCharFormat::AlignSubScript : QTextCharFormat::AlignNormal);
            });
        }
        if (n == "selprotected") { if (!set) return sel_char([&](const QTextCharFormat &f) { return QVariant(f.boolProperty(RT_PROTECTED)); }); bool v = vb_v2bool(val) != 0; return setchar([&](QTextCharFormat &f) { f.setProperty(RT_PROTECTED, v); }); }
        /* the paragraphs of the selection */
        if (n == "selalignment") {
            if (!set) return sel_block([&](const QTextBlock &b) {
                Qt::Alignment al = b.blockFormat().alignment() & Qt::AlignHorizontal_Mask;
                return QVariant(al == Qt::AlignRight ? 1 : al == Qt::AlignHCenter ? 2 : 0); });
            int v = vb_v2l(val); if (vb_errpend) return vb_v_empty();
            if (v < 0 || v > 2) { vb_raise(380); return vb_v_empty(); }
            each_block([&](QTextCursor &, QTextBlockFormat &bf) { bf.setAlignment(v == 1 ? Qt::AlignRight : v == 2 ? Qt::AlignHCenter : Qt::AlignLeft); });
            return vb_v_empty();
        }
        if (n == "selindent" || n == "selhangingindent" || n == "selrightindent") {
            /* RichEdit: SelIndent is where the first line starts, SelHangingIndent where the others start from it */
            if (!set) return sel_block([&](const QTextBlock &b) {
                QTextBlockFormat f = b.blockFormat();
                double v = n == "selindent" ? f.leftMargin() + f.textIndent() : n == "selhangingindent" ? -f.textIndent() : f.rightMargin();
                return QVariant(tw(v)); });
            double v = px_pixels(val); if (vb_errpend) return vb_v_empty();        /* whole pixels (100 -> 105 twips) */
            each_block([&](QTextCursor &, QTextBlockFormat &bf) {
                if (n == "selrightindent") { bf.setRightMargin(v); return; }
                double first = bf.leftMargin() + bf.textIndent(), hang = -bf.textIndent();
                if (n == "selindent") first = v; else hang = v;
                bf.setLeftMargin(first + hang); bf.setTextIndent(-hang);
            });
            return vb_v_empty();
        }
        if (n == "selbullet") {
            if (!set) return sel_block([&](const QTextBlock &b) { return QVariant(b.textList() != nullptr); });
            bool v = vb_v2bool(val) != 0;
            each_block([&](QTextCursor &, QTextBlockFormat &bf) { if (v) bf.setProperty(RT_BULLET, true); else bf.clearProperty(RT_BULLET); });
            return vb_v_empty();
        }
        if (n == "bulletindent") { if (!set) return vb_v_f((float)from_twips_ctl(c, bulletindent)); bulletindent = to_twips_ctl(c, vb_v2d(val)); return vb_v_empty(); }
        if (n == "seltabcount") {
            if (!set) return sel_block([&](const QTextBlock &b) { return QVariant((int)b.blockFormat().tabPositions().size()); });
            int v = vb_v2l(val); if (vb_errpend) return vb_v_empty();
            if (v < 0 || v > 32) { vb_raise(380); return vb_v_empty(); }
            each_block([&](QTextCursor &, QTextBlockFormat &bf) {
                QList<QTextOption::Tab> t = bf.tabPositions();
                while (t.size() > v) t.removeLast();
                while (t.size() < v) t.append(QTextOption::Tab((t.isEmpty() ? 0 : t.last().position) + 720 / RT_TPP, QTextOption::LeftTab));
                bf.setTabPositions(t);
            });
            return vb_v_empty();
        }
        if (n == "seltabs") {
            if (argc < 1 + (set ? 1 : 0)) { vb_raise(450); return vb_v_empty(); }
            int i = vb_v2l(argv[0]); if (vb_errpend) return vb_v_empty();
            if (!set) return sel_block([&](const QTextBlock &b) {
                QList<QTextOption::Tab> t = b.blockFormat().tabPositions();
                return i >= 0 && i < t.size() ? QVariant(tw(t[i].position)) : QVariant(0.0); });
            double v = px_pixels(val);
            each_block([&](QTextCursor &, QTextBlockFormat &bf) {
                QList<QTextOption::Tab> t = bf.tabPositions();
                if (i < 0 || i >= t.size()) return;
                t[i].position = v; bf.setTabPositions(t);
            });
            return vb_v_empty();
        }
        /* methods */
        if (n == "find") {
            QString what = arg_s(argc, argv, 0);
            QString q = te->toPlainText();
            int a = given(argc, argv, 1) ? qpos(vb_v2l(argv[1])) : 0;
            int e = given(argc, argv, 2) ? qpos(vb_v2l(argv[2])) : q.size();
            int opt = given(argc, argv, 3) ? vb_v2l(argv[3]) : 0;
            if (vb_errpend) return vb_v_empty();
            Qt::CaseSensitivity cs = opt & 4 ? Qt::CaseSensitive : Qt::CaseInsensitive;
            for (int k = q.indexOf(what, a, cs); k >= 0 && k + what.size() <= e; k = q.indexOf(what, k + 1, cs)) {
                if (opt & 2) {
                    auto wordch = [](QChar ch) { return ch.isLetterOrNumber() || ch == '_'; };
                    if ((k > 0 && wordch(q[k - 1])) || (k + what.size() < q.size() && wordch(q[k + what.size()]))) continue;
                }
                if (!(opt & 8)) { tc.setPosition(k); tc.setPosition(k + (int)what.size(), QTextCursor::KeepAnchor); te->setTextCursor(tc); }
                return vb_v_l(vbpos(k));
            }
            return vb_v_l(-1);
        }
        if (n == "getlinefromchar") {
            int p = qpos(vb_v2l(argc ? argv[0] : vb_v_l(0)));
            int line = 0;
            doc()->documentLayout();
            for (QTextBlock bl = doc()->begin(); bl.isValid(); bl = bl.next()) {
                QTextLayout *ly = bl.layout();
                int lc = ly ? ly->lineCount() : 0;
                if (p < bl.position() + bl.length()) {
                    if (ly && lc) { QTextLine l = ly->lineForTextPosition(p - bl.position()); if (l.isValid()) line += l.lineNumber(); }
                    return vb_v_l(line);
                }
                line += std::max(1, lc);
            }
            return vb_v_l(std::max(0, line - 1));
        }
        if (n == "span" || n == "upto") {
            QString cs = arg_s(argc, argv, 0);
            bool fwd = given(argc, argv, 1) ? vb_v2bool(argv[1]) != 0 : true;
            bool neg = given(argc, argv, 2) ? vb_v2bool(argv[2]) != 0 : false;
            if (vb_errpend) return vb_v_empty();
            QString q = te->toPlainText();
            int a = tc.selectionStart(), b = tc.selectionEnd();
            if (n == "span") {         /* the selection grows over the characters of the set */
                if (fwd) { while (b < q.size() && in_set(pch(q, b), cs, neg)) b++; tc.setPosition(a); tc.setPosition(b, QTextCursor::KeepAnchor); }
                else { while (a > 0 && in_set(pch(q, a - 1), cs, neg)) a--; tc.setPosition(b); tc.setPosition(a, QTextCursor::KeepAnchor); }
            } else {                   /* the insertion point goes up to the first character of the set */
                int p = fwd ? b : a;
                if (fwd) { while (p < q.size() && !in_set(pch(q, p), cs, neg)) p++; }
                else { while (p > 0 && !in_set(pch(q, p - 1), cs, neg)) p--; }
                tc.setPosition(p);
            }
            te->setTextCursor(tc);
            return vb_v_empty();
        }
        if (n == "loadfile") {
            if (argc < 1) { vb_raise(450); return vb_v_empty(); }
            int type = given(argc, argv, 1) ? vb_v2l(argv[1]) : 0;
            QString path = arg_s(argc, argv, 0);
            if (load_file(path, type, true)) filename = path;
            return vb_v_empty();
        }
        if (n == "savefile") {
            if (argc < 1) { vb_raise(450); return vb_v_empty(); }
            int type = given(argc, argv, 1) ? vb_v2l(argv[1]) : 0;
            QFile f(native(arg_s(argc, argv, 0)));
            if (!f.open(QIODevice::WriteOnly | QIODevice::Truncate)) { vb_raise(75); return vb_v_empty(); }
            if (type == 0) f.write(rtf_write(doc(), 0, doc()->characterCount() - 1, true));
            else f.write(ansi(vbtext(te->toPlainText())));
            return vb_v_empty();
        }
        if (n == "filename") {
            if (!set) return vstr(filename);
            QString v = qs(vb_v2s(val));
            if (load_file(v, 0, true)) filename = v;
            return vb_v_empty();
        }
        if (n == "selprint" || n == "refresh") return vb_v_empty();
        if (n == "locked") { if (!set) return vbool(te->isReadOnly()); te->setReadOnly(vb_v2bool(val) != 0); return vb_v_empty(); }
        if (n == "maxlength") { if (!set) return vb_v_l(maxlength); maxlength = vb_v2l(val); return vb_v_empty(); }
        if (n == "multiline") { if (set) { vb_raise(383); return vb_v_empty(); } return vbool(multiline); }
        if (n == "scrollbars") { if (set) { vb_raise(383); return vb_v_empty(); } return vb_v_i((vb_i16)scrollbars); }
        if (n == "hideselection") { if (!set) return vbool(hidesel); hidesel = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (n == "autoverbmenu") { if (!set) return vbool(autoverb); autoverb = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (n == "rightmargin") { if (!set) return vb_v_f((float)rightmargin); rightmargin = vb_v2d(val); return vb_v_empty(); }
        if (n == "disablenoscroll" || n == "oledragmode" || n == "oledropmode") { if (!set) return vb_v_l(0); return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
    double bulletindent = 0;
};
