/* ocx_mask.inc - the Masked Edit control (MSMASK32.OCX, MSMask.MaskEdBox): a QLineEdit whose typing the
 * control does itself. Included by ocx.inc.
 *
 * As VB6 (oracle tests/dev/f_mask), the text is a string with one character for each position of the mask:
 * the literals, the characters typed, and the PromptChar where nothing is (so a PromptChar that is also a
 * valid character - "1" with "#" - makes those characters empty: Clip and FormattedText leave them out).
 *   Mask: # 9 digit, ? letter, A a letter or digit, & C any character from 32 (not 127), > upper case from here,
 *         < lower case, \ the next character is a literal; every other character is a literal (. , : / too,
 *         whatever the regional settings). MaxLength becomes the number of positions; Mask = "" makes it 64.
 *         A new mask keeps the old text when its first characters fit (two Change events), else empties it.
 *   Text: with a mask it must fit it exactly, else 380 (PromptInclude False: positions may be missing, a space
 *         is empty, and Text gives ClipText). Change on every Text = that fits, even with the same text.
 *   Typing goes in at the caret (the literals are passed over) and pushes what is after it to the right; a
 *         character that does not fit, or one pushed out at the end, gives ValidationError(InvalidText,
 *         StartPosition) and nothing changes. A selection is taken out first (what follows moves left).
 *         Backspace / Delete take out a character the same way. AutoTab: Tab when all positions are filled.
 *   ClipText, FormattedText are read-only (383). FormattedText: Format$(Text, Format) when Format is set,
 *         else the text with spaces for the PromptChar. Out of focus the control shows FormattedText.
 *   SelText = inserts as typing does; then SelStart and SelLength are 0 (as VB6).
 * Bugs of VB6 not copied (oracle f_mask):
 *   - a Text = or SelText = that does not fit raised 380 but fired Change all the same, and a Text longer
 *     than MaxLength made MaxLength grow even when it was refused: here nothing happens to them;
 *   - SelText with ClipMode = mskExcludeLiterals gave all the ClipText from SelStart on, not only the
 *     selection: here it is the characters of the selection without the literals.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

struct OcxMaskEdBox;
class MaskLine : public QLineEdit {
public:
    vbobj self = nullptr;
    explicit MaskLine(QWidget *parent) : QLineEdit(parent) {}
protected:
    void keyPressEvent(QKeyEvent *e) override;
    void focusInEvent(QFocusEvent *e) override;
    void focusOutEvent(QFocusEvent *e) override;
};

struct OcxMaskEdBox : Ocx {
    struct Pos { QChar lit; char kind = 0; int kase = 0; };    /* kind 0: a literal; kase 1 upper, 2 lower */
    QPointer<MaskLine> le;
    QString mask, text, format;
    std::vector<Pos> pos;
    QChar prompt = QChar('_');
    bool promptinclude = true, allowprompt = false, autotab = false, hidesel = true, quiet = false;
    int clipmode = 0, maxlength = 64, appearance = 1, border = 1, mousepointer = 0, oledrag = 0, oledrop = 0;

    bool masked() const { return !pos.empty(); }
    const char *bound_prop() override { return "text"; }  /* DataField: Text */
    bool focusable() override { return true; }
    int n() const { return (int)pos.size(); }

    QWidget *make(QWidget *parent) override
    {
        le = new MaskLine(parent);
        le->self = c->self;
        le->setContextMenuPolicy(Qt::NoContextMenu);
        le->setAcceptDrops(false);
        le->setMaxLength(maxlength);
        vbobj self = c->self;
        /* typing without a mask: the QLineEdit does it */
        QObject::connect(le, &QLineEdit::textEdited, [self](const QString &t) {
            Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxMaskEdBox *>(x->ocx) : nullptr;
            if (!o || o->masked()) return;
            o->text = t;
            o->changed();
        });
        return le;
    }

    /* ------------------------------------------------------------ the mask */
    static std::vector<Pos> parse(const QString &m)
    {
        std::vector<Pos> r;
        int kase = 0;
        for (int i = 0; i < m.size(); i++) {
            QChar ch = m[i];
            Pos p;
            if (ch == '>') { kase = 1; continue; }
            if (ch == '<') { kase = 2; continue; }
            if (ch == '\\') { if (++i >= m.size()) break; p.lit = m[i]; }
            else if (QString("#9?Aa&C").contains(ch)) { p.kind = (char)ch.toLatin1(); p.kase = kase; }
            else p.lit = ch;
            r.push_back(p);
        }
        return r;
    }
    /* a character at a position of the mask: false when it does not fit; *out with its case */
    bool fits(QChar ch, const Pos &p, QChar *out) const
    {
        ushort u = ch.unicode();
        bool ok;
        switch (p.kind) {
        case '#': case '9': ok = u >= '0' && u <= '9'; break;
        case '?': ok = ch.isLetter(); break;
        case 'A': case 'a': ok = ch.isLetterOrNumber(); break;
        default: ok = u >= 32 && u != 127; break;       /* & C */
        }
        if (!ok) return false;
        *out = p.kase == 1 ? ch.toUpper() : p.kase == 2 ? ch.toLower() : ch;
        return true;
    }
    QString empty_text() const
    {
        QString t;
        for (auto &p : pos) t += p.kind ? prompt : p.lit;
        return t;
    }
    /* Text = with a mask: the whole text (PromptInclude) or the first positions (a space is empty) */
    bool parse_text(const QString &s, bool lenient, QString *out) const
    {
        if (lenient ? s.size() > n() : s.size() != n()) return false;
        QString t;
        for (int i = 0; i < n(); i++) {
            const Pos &p = pos[i];
            if (i >= s.size()) { t += p.kind ? prompt : p.lit; continue; }
            QChar ch = s[i], o;
            if (!p.kind) { if (ch != p.lit) return false; t += ch; continue; }
            if (ch == prompt || (lenient && ch == ' ')) { t += prompt; continue; }
            if (!fits(ch, p, &o)) return false;
            t += o;
        }
        *out = t;
        return true;
    }
    QString clip_of(const QString &t, int a, int b) const
    {
        QString r;
        for (int i = a; i < b && i < n() && i < t.size(); i++) if (pos[i].kind && t[i] != prompt) r += t[i];
        return r;
    }
    QString clip() const { return masked() ? clip_of(text, 0, n()) : text; }
    QString formatted() const
    {
        if (!format.isEmpty()) {
            vb_i32 mk = vb_mark();
            QString r = qs(vb_format(vstr(text), vs(format), 1, 1, 1));
            vb_release(mk);
            return r;
        }
        if (!masked()) return text;
        QString t = text;
        for (int i = 0; i < n() && i < t.size(); i++) if (pos[i].kind && t[i] == prompt) t[i] = ' ';
        return t;
    }
    QString text_prop() const { return masked() && !promptinclude ? clip() : text; }

    /* ------------------------------------------------------------ the widget */
    void show_text(int a, int b)
    {
        if (!le) return;
        QSignalBlocker bl(le);
        bool focus = le->hasFocus();
        le->setText(!focus && !format.isEmpty() ? formatted() : text);
        a = std::max(0, std::min(a, (int)le->text().size()));
        b = std::max(0, std::min(b, (int)le->text().size()));
        if (b > a) le->setSelection(a, b - a); else le->setCursorPosition(a);
    }
    void changed()
    {
        if (!quiet) fire_ctl(c, "change");
    }
    void set_maxlength(int m)
    {
        maxlength = m;
        if (!le) return;
        int a = le->selectionStart(), len = (int)le->selectedText().size(), p = le->cursorPosition();
        QSignalBlocker bl(le);
        le->setMaxLength(masked() ? 32767 : m);         /* Qt puts the caret at the end: as it was */
        if (a >= 0 && len) le->setSelection(a, len); else le->setCursorPosition(std::min(p, (int)le->text().size()));
    }

    /* the positions of the mask from i on that take characters */
    std::vector<int> slots_from(int i) const { std::vector<int> r; for (; i < n(); i++) if (pos[i].kind) r.push_back(i); return r; }
    /* the characters of [a, b) taken out, what follows moves to the left: false when one does not fit */
    bool take_out(QString &t, int a, int b) const
    {
        std::vector<int> s = slots_from(a), keep;
        std::vector<QChar> rest;
        for (int i : s) if (i >= b) rest.push_back(t[i]);
        for (size_t k = 0; k < s.size(); k++) {
            QChar ch = k < rest.size() ? rest[k] : prompt, o = ch;
            if (ch != prompt && !fits(ch, pos[s[k]], &o)) return false;
            t[s[k]] = o;
        }
        return true;
    }
    /* ch put in at the position i (a position that takes characters), what is there moves to the right.
     * 0: done; else the InvalidText and the StartPosition of ValidationError */
    bool put_in(QString &t, int i, QChar ch, QString *inv, int *at) const
    {
        QChar o;
        if (!(ch == prompt ? allowprompt : fits(ch, pos[i], &o))) {
            QString r = t.left(i) + ch;
            for (int j = i + 1; j < n(); j++) r += pos[j].kind ? prompt : pos[j].lit;
            *inv = r; *at = i;
            return false;
        }
        if (ch == prompt) o = prompt;
        std::vector<int> s = slots_from(i);
        std::vector<QChar> seq{ o };
        for (int k : s) seq.push_back(t[k]);
        QString r = t;
        bool bad = false; int badat = n();
        for (size_t k = 0; k < s.size(); k++) {
            QChar ch2 = seq[k], o2 = ch2;
            if (k > 0 && ch2 != prompt && !fits(ch2, pos[s[k]], &o2)) { if (!bad) { bad = true; badat = s[k]; } o2 = ch2; }
            r[s[k]] = o2;
        }
        if (seq.back() != prompt) bad = true;           /* pushed out at the end */
        if (bad) { *inv = r; *at = badat; return false; }
        t = r;
        return true;
    }
    void validation_error(const QString &inv, int at)
    {
        vbvar a[2] = { vstr(inv), vb_v_i((vb_i16)at) };
        fire_ctl(c, "validationerror", 2, a);
    }
    bool full() const { for (int i = 0; i < n(); i++) if (pos[i].kind && text[i] == prompt) return false; return true; }
    void after_typing(int caret)
    {
        QPointer<MaskLine> w(le);
        show_text(caret, caret);
        changed();
        if (vb_errpend || !w) return;
        if (autotab && full())
            if (FormState *st = form_of_widget(w)) tab_move(st, false);
    }

    /* typing (MaskLine::keyPressEvent) */
    void type_char(QChar ch)
    {
        int a = le->selectionStart(), b = a + (int)le->selectedText().size(), p = le->cursorPosition();
        QString t = text;
        if (a < 0) a = b = p;
        if (b > a) { if (!take_out(t, a, b)) return; p = a; }
        else if (p < n() && !pos[p].kind && ch == pos[p].lit) {     /* the literal itself: past it */
            show_text(p + 1, p + 1);
            return;
        }
        int q = p;
        while (q < n() && !pos[q].kind) q++;
        QString inv; int at;
        if (q >= n()) { validation_error(t, p); return; }
        if (!put_in(t, q, ch, &inv, &at)) { validation_error(inv, at); return; }
        text = t;
        int caret = q + 1;
        while (caret < n() && !pos[caret].kind) caret++;
        after_typing(caret);
    }
    void erase(bool back)
    {
        int a = le->selectionStart(), b = a + (int)le->selectedText().size(), p = le->cursorPosition();
        QString t = text;
        if (a >= 0 && b > a) {
            if (!take_out(t, a, b)) return;
            text = t;
            after_typing(a);
            return;
        }
        int q = p;
        if (back) {
            q = p - 1;
            while (q >= 0 && !pos[q].kind) q--;
            if (q < 0) return;
        } else {
            while (q < n() && !pos[q].kind) q++;
            if (q >= n()) return;
        }
        if (!take_out(t, q, q + 1)) return;
        text = t;
        after_typing(back ? q : p);
    }
    /* SelText = and Paste: the selection taken out, the characters put in one after another */
    bool insert_text(const QString &s, int a, int b, int *caret)
    {
        QString t = text;
        if (b > a && !take_out(t, a, b)) return false;
        int q = a;
        for (QChar ch : s) {
            while (q < n() && !pos[q].kind) {
                if (ch == pos[q].lit) break;
                q++;
            }
            if (q < n() && !pos[q].kind) { q++; continue; }      /* the literal itself */
            QString inv; int at;
            if (q >= n() || !put_in(t, q, ch, &inv, &at)) return false;
            q++;
        }
        text = t;
        *caret = q;
        return true;
    }

    /* ------------------------------------------------------------ the properties */
    bool set_mask(const QString &m, bool events)
    {
        std::vector<Pos> np = parse(m);
        if ((int)np.size() > 64) { vb_raise(380); return false; }
        if (np.empty() && !masked()) { mask = m; return true; }    /* no mask before and after: no Change */
        QString old = text;
        mask = m; pos = np;
        if (!masked()) {
            set_maxlength(64);
            show_text(0, 0);
            if (events) changed();
            return true;
        }
        set_maxlength(n());
        QString t;
        text = empty_text();
        show_text(0, 0);
        if (events) changed();
        if (parse_text(old.left(n()), false, &t)) {     /* the old text fits: kept */
            text = t;
            show_text(0, 0);
            if (events) changed();
        }
        return true;
    }
    void load(const char *const *p) override
    {
        QString v;
        quiet = true;
        if (prop(p, "PromptChar", &v) && v.size() == 1) prompt = v[0];
        promptinclude = propn(p, "PromptInclude", -1) != 0;
        allowprompt = propn(p, "AllowPrompt", 0) != 0;
        autotab = propn(p, "AutoTab", 0) != 0;
        hidesel = propn(p, "HideSelection", -1) != 0;
        clipmode = (int)propn(p, "ClipMode", 0);
        appearance = (int)propn(p, "Appearance", 1);
        border = (int)propn(p, "BorderStyle", 1);
        mousepointer = (int)propn(p, "MousePointer", 0);
        oledrag = (int)propn(p, "OLEDragMode", 0);
        oledrop = (int)propn(p, "OLEDropMode", 0);
        if (prop(p, "Format", &v)) format = v;
        if (prop(p, "Text", &v)) text = v;
        if (prop(p, "Mask", &v)) set_mask(v, false);
        int ml = (int)propn(p, "MaxLength", masked() ? n() : 64);
        if (ml >= 1 && ml <= 64) {
            if (masked() && ml < n()) { pos.resize(ml); text = text.left(ml); }
            set_maxlength(ml);
        }
        if (le) le->setFrame(border != 0);
        show_text(0, 0);
        quiet = false;
    }

    vbvar member(const std::string &nm, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (!le) { *found = 0; return vb_v_empty(); }
        if (nm == "text" || nm == "defaulttext") {
            if (!set) return vstr(text_prop());
            QString s = qs(vb_v2s(val));
            if (vb_errpend) return vb_v_empty();
            if (masked()) {
                QString t;
                if (!parse_text(s, !promptinclude, &t)) { vb_raise(380); return vb_v_empty(); }
                text = t;
            } else {
                text = s;
                if (s.size() > maxlength) set_maxlength((int)s.size());   /* as VB6: MaxLength grows */
            }
            show_text(0, 0);
            changed();
            return vb_v_empty();
        }
        if (nm == "cliptext" || nm == "formattedtext") {
            if (set) { vb_raise_full(383, vs(QString("MaskEdBox")), vs(QString("Property is read-only"))); return vb_v_empty(); }
            return vstr(nm == "cliptext" ? clip() : formatted());
        }
        if (nm == "mask") {
            if (!set) return vstr(mask);
            QString m = qs(vb_v2s(val));
            if (vb_errpend) return vb_v_empty();
            set_mask(m, true);
            return vb_v_empty();
        }
        if (nm == "promptchar") {
            if (!set) return vstr(QString(prompt));
            QString v = qs(vb_v2s(val));
            if (vb_errpend) return vb_v_empty();
            if (v.size() != 1) { vb_raise(380); return vb_v_empty(); }
            QChar old = prompt;
            prompt = v[0];
            for (int i = 0; i < n() && i < text.size(); i++) if (pos[i].kind && text[i] == old) text[i] = prompt;
            show_text(le->cursorPosition(), le->cursorPosition());
            changed();
            return vb_v_empty();
        }
        if (nm == "format") {
            if (!set) return vstr(format);
            QString v = qs(vb_v2s(val));
            if (vb_errpend) return vb_v_empty();
            format = v;
            show_text(le->cursorPosition(), le->cursorPosition());
            changed();
            return vb_v_empty();
        }
        if (nm == "maxlength") {
            if (!set) return vb_v_i((vb_i16)maxlength);
            int v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 1 || v > 64) { vb_raise(380); return vb_v_empty(); }
            if (masked()) {
                if (v < n()) {                  /* the mask is cut there */
                    pos.resize(v);
                    text = text.left(v);
                    set_maxlength(v);
                    show_text(0, 0);
                    changed();
                } else set_maxlength(v);
            } else {
                set_maxlength(v);
                if (text.size() > v) { text = text.left(v); show_text(0, 0); changed(); }
            }
            return vb_v_empty();
        }
        struct { const char *name; bool *p; } bools[] = { { "promptinclude", &promptinclude }, { "allowprompt", &allowprompt },
                                                          { "autotab", &autotab }, { "hideselection", &hidesel } };
        for (auto &e : bools) if (nm == e.name) {
            if (!set) return vbool(*e.p);
            bool v = vb_v2bool(val) != 0;
            if (vb_errpend) return vb_v_empty();
            *e.p = v;
            return vb_v_empty();
        }
        struct { const char *name; int *p; int lo, hi; } ints[] = { { "clipmode", &clipmode, 0, 1 }, { "appearance", &appearance, 0, 1 },
                                                                    { "borderstyle", &border, 0, 1 }, { "oledragmode", &oledrag, 0, 1 },
                                                                    { "oledropmode", &oledrop, 0, 2 } };
        for (auto &e : ints) if (nm == e.name) {
            if (!set) return vb_v_l(*e.p);
            int v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < e.lo || v > e.hi) { vb_raise(380); return vb_v_empty(); }
            *e.p = v;
            if (e.p == &border) le->setFrame(border != 0);
            return vb_v_empty();
        }
        if (nm == "selstart" || nm == "sellength" || nm == "seltext") {
            int a = le->selectionStart(), len = (int)le->selectedText().size();
            if (a < 0) { a = le->cursorPosition(); len = 0; }
            if (!set) {
                if (nm == "selstart") return vb_v_l(a);
                if (nm == "sellength") return vb_v_l(len);
                if (masked() && clipmode == 1) return vstr(clip_of(text, a, a + len));
                return vstr(text.mid(a, len));
            }
            if (nm == "seltext") {
                QString s = qs(vb_v2s(val));
                if (vb_errpend) return vb_v_empty();
                if (masked()) {
                    int caret;
                    if (!insert_text(s, a, a + len, &caret)) return vb_v_empty();
                    show_text(0, 0);            /* VB6: SelStart and SelLength are 0 after it */
                } else {
                    text = text.left(a) + s + text.mid(a + len);
                    show_text(a + (int)s.size(), a + (int)s.size());
                }
                changed();
                return vb_v_empty();
            }
            int v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0) { vb_raise(380); return vb_v_empty(); }
            int tl = (int)text.size();
            if (nm == "selstart") { v = std::min(v, tl); le->setCursorPosition(v); }
            else { int e2 = std::min(a + v, tl); if (e2 > a) le->setSelection(a, e2 - a); else le->setCursorPosition(a); }
            return vb_v_empty();
        }
        if (nm == "mousepointer") { if (!set) return vb_v_l(mousepointer); mousepointer = vb_v2l(val); *found = 0; return vb_v_empty(); }
        if (nm == "aboutbox" || nm == "oledrag") return vb_v_empty();
        if (nm == "refresh") { le->update(); return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
};

void MaskLine::keyPressEvent(QKeyEvent *e)
{
    Ctl *x = C(self);
    auto *o = x ? dynamic_cast<OcxMaskEdBox *>(x->ocx) : nullptr;
    if (!o || !o->masked() || isReadOnly()) { QLineEdit::keyPressEvent(e); return; }
    if (e->matches(QKeySequence::Undo) || e->matches(QKeySequence::Redo) || e->matches(QKeySequence::DeleteEndOfWord) ||
        e->matches(QKeySequence::DeleteStartOfWord) || e->matches(QKeySequence::DeleteEndOfLine)) return;
    if (e->matches(QKeySequence::Copy) || e->matches(QKeySequence::Cut)) {
        int a = selectionStart(), len = (int)selectedText().size();
        if (a < 0 || !len) return;
        QApplication::clipboard()->setText(o->clipmode == 1 ? o->clip_of(o->text, a, a + len) : o->text.mid(a, len));
        if (e->matches(QKeySequence::Cut)) o->erase(true);
        return;
    }
    if (e->matches(QKeySequence::Paste)) {
        int a = selectionStart(), len = (int)selectedText().size(), caret;
        if (a < 0) { a = cursorPosition(); len = 0; }
        if (o->insert_text(QApplication::clipboard()->text(), a, a + len, &caret)) o->after_typing(caret);
        return;
    }
    if (e->key() == Qt::Key_Backspace) { o->erase(true); return; }
    if (e->key() == Qt::Key_Delete) { o->erase(false); return; }
    QString t = e->text();
    if (t.size() == 1 && t[0].unicode() >= 32 && t[0].unicode() != 127 && !(e->modifiers() & (Qt::ControlModifier | Qt::AltModifier))) {
        o->type_char(t[0]);
        return;
    }
    if (t.size() == 1 && t[0].unicode() < 32 && e->key() != Qt::Key_Return && e->key() != Qt::Key_Enter && e->key() != Qt::Key_Escape
        && e->key() != Qt::Key_Tab) return;     /* other control characters: nothing */
    QLineEdit::keyPressEvent(e);
}
void MaskLine::focusInEvent(QFocusEvent *e)
{
    Ctl *x = C(self);
    auto *o = x ? dynamic_cast<OcxMaskEdBox *>(x->ocx) : nullptr;
    QLineEdit::focusInEvent(e);
    if (o && !o->format.isEmpty()) { int p = cursorPosition(); o->show_text(p, p); }
}
void MaskLine::focusOutEvent(QFocusEvent *e)
{
    Ctl *x = C(self);
    auto *o = x ? dynamic_cast<OcxMaskEdBox *>(x->ocx) : nullptr;
    QLineEdit::focusOutEvent(e);
    if (o && !o->format.isEmpty()) o->show_text(0, 0);
}

static Ocx *new_maskedbox() { return new OcxMaskEdBox; }
