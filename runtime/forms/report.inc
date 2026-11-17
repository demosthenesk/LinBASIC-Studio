/* report.inc - the DataReport of VB6 (MSDataReportLib): the report a DataReport designer (.dsr)
 * describes, its sections and controls (RptLabel, RptTextBox, RptImage, RptLine, RptShape,
 * RptFunction), its data (the DataSource / DataMember: a recordset of the DataEnvironment, with
 * the chapters of its hierarchy for the group sections), the pages, and what the program does with
 * them: Show (the preview window), PrintReport (a PDF file, as the Printer object), ExportReport
 * (HTML, text, Unicode HTML and text through the ExportFormats and their templates).
 * The compiler makes a class of the .dsr (compiler/lbvc/designers.py) that keeps this object
 * (vbf_report_new) and sends its events to the DataReport_ procedures of the code.
 * Members, errors, the default export formats and templates as VB6 has them (VB6 oracles
 * tests/db/d_rpt1, d_rpt2). The layout of the pages is LinBASIC Studio's own: VB6 measures the
 * text with the fonts of Windows; the exports have the same structure, not the same bytes.
 * Included by vbforms.cpp. Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

#include <QToolBar>
extern "C" char *vb_stoansi(vbs s, int *n);
#include <QScrollArea>

enum { SEC_RH = 1, SEC_PH = 2, SEC_GH = 3, SEC_DETAIL = 4, SEC_GF = 5, SEC_PF = 7, SEC_RF = 8 };

struct RFont { QString name = "Arial"; double size = 8.25; bool bold = false, italic = false, underline = false, strike = false; };
struct RptCtl {
    QString type, name, caption, datafield, datamember, fmt, picture;
    int fmttype = 0;
    double left = 0, top = 0, width = 0, height = 0;
    int alignment = 0, backstyle = 0, borderstyle = 0, lineslant = 0, shape = 0, sizemode = 0, functype = 0, picalign = 8;
    long backcolor = 0xFFFFFF, forecolor = 0, bordercolor = 0;
    bool visible = true, cangrow = false, rtl = false;
    RFont font;
    QImage img;                                 /* RptImage: its Picture (from the .dsx) */
    vbobj obj = nullptr;                        /* the VB object of the control, made when first asked */
    struct RptSec *sec = nullptr;
};
struct RptSec {
    int code = SEC_DETAIL;
    QString name;
    double height = 0;
    bool keep = false, visible = true;
    int forcebreak = 0;
    std::vector<RptCtl *> ctls;
    vbobj obj = nullptr, ctlsobj = nullptr;
    struct Report *rep = nullptr;
};
struct RptFormat { QString key, tmpl, filestring, filter; int type = 0; vbobj obj = nullptr; };
/* what the pages are made of */
struct RptItem {
    int kind = 0;                               /* 0 text, 1 line, 2 shape, 3 image */
    double x = 0, y = 0, w = 0, h = 0;
    QString text;
    RFont font;
    int align = 0, slant = 0, shape = 0, backstyle = 0, borderstyle = 1;
    long fore = 0, back = 0xFFFFFF, border = 0;
    QImage img;
    int sizemode = 0, picalign = 8;
};
struct RptPage { std::vector<RptItem> items; };

struct Report {
    vbobj self = nullptr, owner = nullptr;      /* owner: the object of the .dsr class (not counted) */
    QString caption, title, datamember;
    double lm = 1440, rm = 1440, tm = 1440, bm = 1440, width = 0, gridx = 10, gridy = 10;
    int orientation = 0;
    RFont font;
    vbobj datasource = nullptr;
    std::vector<RptSec *> secs;
    std::vector<RptFormat *> formats;
    vbobj secsobj = nullptr, fmtsobj = nullptr;
    /* the window */
    QWidget *win = nullptr;
    double wl = -1, wt = -1, ww = -1, wh = -1;
    int windowstate = 0;
    std::vector<RptPage> pages;
    int zoom = 100, curpage = 0;
};

static vbvar rep_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
static vbvar rsecs_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
static vbvar rsec_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
static vbvar rctls_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
static vbvar rfmts_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
static vbvar rfmt_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
static vbvar rptctl_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
static void rep_clear(vbobj);
static const vb_class CLS_REPORT = { "DataReport", (vb_i32)sizeof(vbobjh), nullptr, nullptr, rep_clear, nullptr, 0, nullptr,
                                     nullptr, nullptr, nullptr, rep_member, nullptr };
static const vb_class CLS_RSECTIONS = { "Sections", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item",
                                        nullptr, nullptr, nullptr, rsecs_member, nullptr };
static const vb_class CLS_RSECTION = { "Section", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr,
                                       nullptr, nullptr, nullptr, rsec_member, nullptr };
static const vb_class CLS_RCONTROLS = { "Controls", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item",
                                        nullptr, nullptr, nullptr, rctls_member, nullptr };
static const vb_class CLS_RFORMATS = { "ExportFormats", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item",
                                       nullptr, nullptr, nullptr, rfmts_member, nullptr };
static const vb_class CLS_RFORMAT = { "ExportFormat", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr,
                                      nullptr, nullptr, nullptr, rfmt_member, nullptr };
static vb_class CLS_RLABEL = { "RptLabel", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr, nullptr, nullptr, nullptr, rptctl_member, nullptr };
static vb_class CLS_RTEXTBOX = { "RptTextBox", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr, nullptr, nullptr, nullptr, rptctl_member, nullptr };
static vb_class CLS_RLINE = { "RptLine", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr, nullptr, nullptr, nullptr, rptctl_member, nullptr };
static vb_class CLS_RSHAPE = { "RptShape", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr, nullptr, nullptr, nullptr, rptctl_member, nullptr };
static vb_class CLS_RIMAGE = { "RptImage", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr, nullptr, nullptr, nullptr, rptctl_member, nullptr };
static vb_class CLS_RFUNCTION = { "RptFunction", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr, nullptr, nullptr, nullptr, rptctl_member, nullptr };

static Report *REP(vbobj o) { return o ? (Report *)vb_oid(o)->ext : nullptr; }
static void rpt_raise(vb_i32 n, const char *src, const QString &msg)
{
    vb_i32 m = vb_mark();
    vb_raise_full(n, vb_sfromutf8(src), vs(msg));
    vb_release(m);
}

/* ---------------------------------------------------------------- the definition the compiler wrote */
static std::map<std::string, QString> rpt_fields(const QString &line)
{
    std::map<std::string, QString> m;
    QStringList parts = line.split(QChar(0x1f));
    for (int i = 1; i < parts.size(); i++) {
        int eq = parts[i].indexOf('=');
        if (eq < 0) continue;
        m[parts[i].left(eq).toLower().toStdString()] = parts[i].mid(eq + 1).replace(QChar(0x1e), "\r\n");
    }
    return m;
}
static double num(const std::map<std::string, QString> &m, const char *k, double d)
{
    auto it = m.find(k);
    if (it == m.end() || it->second.isEmpty()) return d;
    bool ok; double v = it->second.toDouble(&ok);
    return ok ? v : d;
}
static QString str(const std::map<std::string, QString> &m, const char *k)
{
    auto it = m.find(k);
    return it == m.end() ? QString() : it->second;
}
static RFont rfont(const std::map<std::string, QString> &m, const RFont &d)
{
    RFont f = d;
    if (m.count("font.name")) f.name = str(m, "font.name");
    f.size = num(m, "font.size", f.size);
    if (m.count("font.weight")) f.bold = num(m, "font.weight", 400) >= 600;
    if (m.count("font.italic")) f.italic = num(m, "font.italic", 0) != 0;
    if (m.count("font.underline")) f.underline = num(m, "font.underline", 0) != 0;
    if (m.count("font.strikethrough")) f.strike = num(m, "font.strikethrough", 0) != 0;
    return f;
}
static void rpt_default_formats(Report *r)
{
    static const char *html = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2 Final//EN\">\r\n<HTML>\r\n<HEAD>\r\n%1<TITLE><!--MSDBRPT_Template_Title--></TITLE>\r\n</HEAD>\r\n<BODY>\r\n\r\n<!--MSDBRPT_Template_Body-->\r\n\r\n</BODY>\r\n</HTML>\r\n";
    auto add = [&](const char *key, int type, const char *fs, const char *filter, const QString &t) {
        RptFormat *f = new RptFormat; f->key = key; f->type = type; f->filestring = fs; f->filter = filter; f->tmpl = t;
        r->formats.push_back(f);
    };
    add("key_def_HTML", 0, "HTML (*.htm; *.html)", "*.htm;*.html", QString(html).arg(""));
    add("key_def_UnicodeHTML_UTF8", 3, "Unicode HTML (UTF-8) (*.htm; *.html)", "*.htm;*.html",
        QString(html).arg("<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html;charset=utf-8\">\r\n"));
    add("key_def_Text", 1, "Text (*.txt)", "*.txt", "<!--MSDBRPT_Template_Body-->");
    add("key_def_UnicodeText", 2, "Unicode Text (*.txt)", "*.txt", "<!--MSDBRPT_Template_Body-->");
}

extern "C" vbobj vbf_report_new(vbobj owner, vbs def)
{
    ensure_app();
    vbobj o = vb_onew(&CLS_REPORT);
    if (!o) return nullptr;
    Report *r = new Report;
    o->ext = r;
    r->self = o; r->owner = vb_oid(owner);
    RptSec *sec = nullptr;
    for (const QString &line : qs(def).split('\n', Qt::SkipEmptyParts)) {
        auto m = rpt_fields(line);
        QChar k = line.isEmpty() ? QChar() : line[0];
        if (k == 'R') {
            r->caption = str(m, "caption"); r->title = str(m, "title"); r->datamember = str(m, "datamember");
            r->lm = num(m, "leftmargin", 1440); r->rm = num(m, "rightmargin", 1440);
            r->tm = num(m, "topmargin", 1440); r->bm = num(m, "bottommargin", 1440);
            r->width = num(m, "reportwidth", 0); r->gridx = num(m, "gridx", 10); r->gridy = num(m, "gridy", 10);
            r->font = rfont(m, RFont());
            r->ww = num(m, "clientwidth", -1); r->wh = num(m, "clientheight", -1);
            r->wl = num(m, "clientleft", -1); r->wt = num(m, "clienttop", -1);
        } else if (k == 'S') {
            sec = new RptSec; sec->rep = r;
            sec->code = (int)num(m, "code", SEC_DETAIL); sec->name = str(m, "name"); sec->height = num(m, "height", 0);
            sec->keep = num(m, "keeptogether", 0) != 0; sec->forcebreak = (int)num(m, "forcepagebreak", 0);
            sec->visible = num(m, "visible", -1) != 0;
            r->secs.push_back(sec);
        } else if (k == 'C' && sec) {
            RptCtl *c = new RptCtl; c->sec = sec;
            c->type = str(m, "type"); c->name = str(m, "name"); c->caption = str(m, "caption");
            c->datafield = str(m, "datafield"); c->datamember = str(m, "datamember");
            c->fmt = str(m, "dataformat.format"); c->fmttype = (int)num(m, "dataformat.type", 0);
            c->left = num(m, "left", 0); c->top = num(m, "top", 0); c->width = num(m, "width", 0); c->height = num(m, "height", 0);
            c->alignment = (int)num(m, "alignment", 0); c->backstyle = (int)num(m, "backstyle", 0);
            c->borderstyle = (int)num(m, "borderstyle", c->type == "RptLine" || c->type == "RptShape" ? 1 : 0);
            c->lineslant = (int)num(m, "lineslant", 0); c->shape = (int)num(m, "shape", 0);
            c->sizemode = (int)num(m, "sizemode", 0); c->functype = (int)num(m, "functiontype", 0);
            c->picalign = (int)num(m, "picturealignment", 8);
            c->backcolor = (long)num(m, "backcolor", 0xFFFFFF); c->forecolor = (long)num(m, "forecolor", 0);
            c->bordercolor = (long)num(m, "bordercolor", 0);
            c->visible = num(m, "visible", -1) != 0; c->cangrow = num(m, "cangrow", 0) != 0;
            c->font = rfont(m, r->font);
            sec->ctls.push_back(c);
        }
    }
    rpt_default_formats(r);
    return o;
}

/* ---------------------------------------------------------------- the data */
static vbvar lget(vbobj o, const char *name, vb_i32 argc = 0, vbvar *argv = nullptr)
{
    return vb_late_o(o, name, VBM_GET, argc, argv);
}
static bool rs_eof(vbobj rs) { vbvar v = lget(rs, "eof"); return vb_errpend || vb_v_truth(v); }
static vbobj rpt_rs(Report *r, QString member)
{
    vbobj src = r->datasource ? vb_oid(r->datasource) : nullptr;
    if (!src) return nullptr;
    if (!strcmp(src->cls->name, "Recordset")) return src;
    vb_i32 save = vb_errpend;
    vb_i32 m = vb_mark();
    vbvar a[2] = { vb_v_s(vs(member)), vb_v_obj(nullptr) };
    vb_late_o(src, "lb_getdatamember", VBM_METHOD, 2, a);
    vb_release(m);
    if (vb_errpend && !save) { vb_err_clear_pending(); return nullptr; }
    return a[1].vt == VT_DISPATCH && a[1].u.o ? vb_oid(a[1].u.o) : nullptr;
}
static vbvar rs_field(vbobj rs, const QString &name)
{
    vb_i32 save = vb_errpend;
    vbvar k = vb_v_s(vs(name));
    vbvar f = lget(rs, "fields", 1, &k);
    if (vb_errpend && !save) { vb_err_clear_pending(); return vb_v_null(); }
    vbvar v = vb_late(f, "value", VBM_GET, 0, nullptr);
    if (vb_errpend && !save) { vb_err_clear_pending(); return vb_v_null(); }
    return v;
}
/* the chapter of the current row that leads to the next level: the one of that name, or the first */
static vbobj rs_chapter(vbobj rs, const QString &name)
{
    vb_i32 save = vb_errpend;
    vbvar flds = lget(rs, "fields");
    vbvar cnt = vb_late(flds, "count", VBM_GET, 0, nullptr);
    int n = vb_errpend ? 0 : vb_v2l(cnt);
    vbobj first = nullptr;
    for (int i = 0; i < n && !vb_errpend; i++) {
        vbvar ix = vb_v_l(i);
        vbvar f = vb_late(flds, "item", VBM_GET, 1, &ix);
        if (vb_v2l(vb_late(f, "type", VBM_GET, 0, nullptr)) != 136) continue;
        QString fn = qs(vb_v2s(vb_late(f, "name", VBM_GET, 0, nullptr)));
        vbvar v = vb_late(f, "value", VBM_GET, 0, nullptr);
        vbobj o = v.vt == VT_DISPATCH ? v.u.o : nullptr;
        if (!o) continue;
        if (!name.isEmpty() && !fn.compare(name, Qt::CaseInsensitive)) { if (first) vb_orelease(first); vb_oaddref(o); return o; }
        if (!first) { first = o; vb_oaddref(o); }
    }
    if (vb_errpend && !save) vb_err_clear_pending();
    return first;
}

/* the text of a value as a control shows it (its DataFormat) */
static QString rpt_text(const RptCtl *c, vbvar v)
{
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) return QString();
    vb_i32 save = vb_errpend, m = vb_mark();
    QString s;
    if (c->fmttype == 1 && !c->fmt.isEmpty()) s = qs(vb_format(v, vs(c->fmt), 1, 1, 1));
    else s = qs(vb_v2s(v));
    vb_release(m);
    if (vb_errpend && !save) { vb_err_clear_pending(); return QString(); }
    return s;
}

/* the functions of the footers: the values of their field in their group */
struct RFunc {
    RptCtl *c; int level;                       /* the level of the rows it reads */
    int n = 0, nv = 0; double sum = 0, sum2 = 0;
    vbvar mn = vb_v_null(), mx = vb_v_null(), total = vb_v_null();
    void reset() { n = nv = 0; sum = sum2 = 0; vb_vfree(&mn); vb_vfree(&mx); vb_vfree(&total); mn = mx = total = vb_v_null(); }
    void add(vbvar v)
    {
        n++;
        if (v.vt == VT_NULL || v.vt == VT_EMPTY) return;
        nv++;
        double d = vb_v2d(v); sum += d; sum2 += d * d;
        vb_i32 save = vb_errpend;
        vbvar t = total.vt == VT_NULL ? v : vb_v_add(total, v);
        if (vb_errpend && !save) vb_err_clear_pending(); else vb_vset(&total, t);
        if (mn.vt == VT_NULL || vb_v_truth(vb_v_cmp(v, mn, 2, 0))) vb_vset(&mn, v);
        if (mx.vt == VT_NULL || vb_v_truth(vb_v_cmp(v, mx, 3, 0))) vb_vset(&mx, v);
        if (vb_errpend && !save) vb_err_clear_pending();
    }
    vbvar value() const
    {
        switch (c->functype) {
        case 0: return nv ? total : vb_v_null();                         /* rptFuncSum */
        case 1: return nv ? vb_v_d(sum / nv) : vb_v_null();             /* Ave */
        case 2: return mn;
        case 3: return mx;
        case 4: return vb_v_l(n);                                        /* RCnt: the rows */
        case 5: return vb_v_l(nv);                                       /* VCnt: the values */
        case 6: case 7: {
            if (nv < 2) return vb_v_null();
            double sd = std::sqrt((sum2 - sum * sum / nv) / (nv - 1));
            return vb_v_d(c->functype == 6 ? sd : sd / std::sqrt((double)nv));
        }
        }
        return vb_v_null();
    }
};

/* ---------------------------------------------------------------- the pages */
struct Layout {
    Report *r;
    double pw = 12240, ph = 15840;              /* the paper: Letter (the printer of VB6's test machine) */
    double y = 0, bottom = 0;
    std::vector<RptPage> pages;
    std::vector<RptSec *> gh, gf;               /* the group headers (outer first) and footers (outer first) */
    RptSec *rh = nullptr, *ph_ = nullptr, *detail = nullptr, *pf = nullptr, *rf = nullptr;
    std::vector<vbobj> cur;                     /* the recordset of each level now */
    std::vector<RFunc> funcs;
    QString err;
    int errnum = 0;

    RptPage &page() { return pages.back(); }
    int level_of(const QString &member)
    {
        /* a DataMember: the level of the section of that name (Customers_Header -> Customers) */
        for (size_t i = 0; i < gh.size(); i++) if (!gh[i]->name.compare(member + "_Header", Qt::CaseInsensitive)) return (int)i;
        if (detail && !detail->name.compare(member + "_Detail", Qt::CaseInsensitive)) return (int)gh.size();
        for (size_t i = 0; i < gh.size(); i++) if (gh[i]->name.startsWith(member + "_", Qt::CaseInsensitive)) return (int)i;
        if (detail && detail->name.startsWith(member + "_", Qt::CaseInsensitive)) return (int)gh.size();
        if (!member.compare(r->datamember, Qt::CaseInsensitive)) return 0;
        return (int)gh.size();
    }
    void place(RptSec *s, bool fill)
    {
        if (!s->visible) return;
        for (RptCtl *c : s->ctls) {
            if (!c->visible) continue;
            RptItem it;
            it.x = r->lm + c->left; it.y = y + c->top; it.w = c->width; it.h = c->height;
            it.font = c->font; it.align = c->alignment; it.fore = c->forecolor; it.back = c->backcolor;
            it.backstyle = c->backstyle; it.borderstyle = c->borderstyle; it.border = c->bordercolor;
            if (c->type == "RptLine") { it.kind = 1; it.slant = c->lineslant; }
            else if (c->type == "RptShape") { it.kind = 2; it.shape = c->shape; }
            else if (c->type == "RptImage") { it.kind = 3; it.img = c->img; it.sizemode = c->sizemode; it.picalign = c->picalign; }
            else if (c->type == "RptLabel") it.text = c->caption;
            else if (c->type == "RptTextBox") {
                int lv = level_of(c->datamember.isEmpty() ? r->datamember : c->datamember);
                if (fill && lv < (int)cur.size() && cur[lv]) it.text = rpt_text(c, rs_field(cur[lv], c->datafield));
                it.kind = 0;
            } else if (c->type == "RptFunction") {
                for (RFunc &f : funcs) if (f.c == c) {
                    vbvar v = f.value();
                    /* without a DataFormat the number of a function is not regional (VB6 oracle d_de2: 121.5) */
                    if ((c->fmttype != 1 || c->fmt.isEmpty()) && v.vt != VT_NULL && v.vt != VT_EMPTY && v.vt != VT_BSTR && v.vt != VT_DATE)
                        it.text = QString::number(vb_v2d(v), 'g', 15);
                    else it.text = rpt_text(c, v);
                    break;
                }
            }
            page().items.push_back(it);
        }
        y += s->height;
    }
    void new_page()
    {
        pages.push_back(RptPage());
        y = r->tm;
        bottom = ph - r->bm - (pf && pf->visible ? pf->height : 0);
        if (ph_) place(ph_, false);
    }
    void end_page()
    {
        if (pf && pf->visible) { double keep = y; y = ph - r->bm - pf->height; place(pf, false); y = keep; }
    }
    void put(RptSec *s, bool fill)
    {
        if (!s || !s->visible) return;
        if ((s->forcebreak & 1) && y > r->tm + (ph_ ? ph_->height : 0)) { end_page(); new_page(); }
        if (y + s->height > bottom + 0.5 && y > r->tm + (ph_ ? ph_->height : 0) + 0.5) { end_page(); new_page(); }
        place(s, fill);
        if (s->forcebreak & 2) { end_page(); new_page(); }
    }
    void feed(int level)
    {
        for (RFunc &f : funcs)
            if (f.level == level && cur[level]) f.add(rs_field(cur[level], f.c->datafield));
    }
    void walk(int level, vbobj rs)
    {
        if (!rs) return;
        vb_late_o(rs, "movefirst", VBM_METHOD, 0, nullptr);
        if (vb_errpend) vb_err_clear_pending();
        int guard = 0;
        while (!rs_eof(rs) && guard++ < 1000000) {
            vb_i32 mk = vb_mark();              /* the temporaries of a row go with it */
            cur[level] = rs;
            if (level < (int)gh.size()) {
                for (RFunc &f : funcs) if (f.c->sec == gf[level]) f.reset();
                feed(level);
                put(gh[level], true);
                QString next = level + 1 < (int)gh.size() ? gh[level + 1]->name : detail ? detail->name : QString();
                int us = next.lastIndexOf('_');
                vbobj ch = rs_chapter(rs, us > 0 ? next.left(us) : QString());
                if (ch) { walk(level + 1, ch); vb_orelease(ch); }
                cur[level] = rs;
                put(gf[level], true);
            } else {
                feed(level);
                put(detail, true);
            }
            vb_late_o(rs, "movenext", VBM_METHOD, 0, nullptr);
            vb_release(mk);
            if (vb_errpend) { vb_err_clear_pending(); break; }
        }
    }
    void run()
    {
        if (r->orientation == 2) std::swap(pw, ph);
        std::vector<RptSec *> headers, footers;
        for (RptSec *s : r->secs) {
            switch (s->code) {
            case SEC_RH: rh = s; break;
            case SEC_PH: ph_ = s; break;
            case SEC_GH: headers.push_back(s); break;
            case SEC_DETAIL: detail = s; break;
            case SEC_GF: footers.push_back(s); break;
            case SEC_PF: pf = s; break;
            case SEC_RF: rf = s; break;
            }
        }
        gh = headers;
        gf.assign(gh.size(), nullptr);          /* the footers are written inner first */
        for (size_t i = 0; i < footers.size() && i < gh.size(); i++) gf[gh.size() - 1 - i] = footers[i];
        cur.assign(gh.size() + 1, nullptr);
        for (RptSec *s : r->secs)
            for (RptCtl *c : s->ctls)
                if (c->type == "RptFunction") {
                    RFunc f; f.c = c; f.level = level_of(c->datamember.isEmpty() ? r->datamember : c->datamember);
                    funcs.push_back(f);
                }
        pages.push_back(RptPage());
        y = r->tm;
        bottom = ph - r->bm - (pf && pf->visible ? pf->height : 0);
        if (rh) place(rh, false);
        if (rh && (rh->forcebreak & 2)) { end_page(); pages.push_back(RptPage()); y = r->tm; }
        if (ph_) place(ph_, false);
        vbobj rs = rpt_rs(r, r->datamember);
        if (!rs && r->datasource) { errnum = 8577; err = "Failed getting Rowset(s) from current data source"; }
        walk(0, rs);
        if (rf) put(rf, true);
        end_page();
        /* %p %P %d %D %t %T %i: the page numbers, the date, the time, the title */
        int np = (int)pages.size();
        QDateTime now = QDateTime::currentDateTime();
        vb_i32 m = vb_mark();
        QString d = qs(vb_format(vb_v_date(vb_now()), vs("Short Date"), 1, 1, 1)), D = qs(vb_format(vb_v_date(vb_now()), vs("Long Date"), 1, 1, 1));
        QString t = qs(vb_format(vb_v_date(vb_now()), vs("Short Time"), 1, 1, 1)), T = qs(vb_format(vb_v_date(vb_now()), vs("Long Time"), 1, 1, 1));
        vb_release(m);
        (void)now;
        for (int p = 0; p < np; p++)
            for (RptItem &it : pages[p].items) {
                if (!it.text.contains('%')) continue;
                QString s, &x = it.text;
                for (int i = 0; i < x.size(); i++) {
                    if (x[i] != '%' || i + 1 >= x.size()) { s += x[i]; continue; }
                    QChar c = x[++i];
                    if (c == 'p') s += QString::number(p + 1);
                    else if (c == 'P') s += QString::number(np);
                    else if (c == 'd') s += d;
                    else if (c == 'D') s += D;
                    else if (c == 't') s += t;
                    else if (c == 'T') s += T;
                    else if (c == 'i') s += r->title;
                    else if (c == '%') s += '%';
                    else { s += '%'; s += c; }
                }
                x = s;
            }
        for (RFunc &f : funcs) f.reset();
    }
};
static bool rpt_layout(Report *r)
{
    Layout L; L.r = r;
    L.run();
    r->pages = L.pages;
    if (L.errnum) { rpt_raise(L.errnum, "DataReport", L.err); return false; }
    return true;
}

/* ---------------------------------------------------------------- the exports */
static int html_size(double pt)
{
    return pt <= 8.5 ? 1 : pt <= 10.5 ? 2 : pt <= 12.5 ? 3 : pt <= 14.5 ? 4 : pt <= 18.5 ? 5 : pt <= 24.5 ? 6 : 7;
}
static QString html_escape(const QString &s)
{
    QString r;
    for (QChar c : s) {
        if (c == '<') r += "&lt;"; else if (c == '>') r += "&gt;"; else if (c == '&') r += "&amp;"; else if (c == '"') r += "&quot;";
        else r += c;
    }
    return r;
}
/* the rows of a page: the things that start at one height; a line (no height) has a row of its own,
 * before the row of the texts it is drawn over (VB6 oracle d_rpt1: the line over a subtotal) */
struct RRow { double y, h; std::vector<const RptItem *> items, spanned; };
static std::vector<RRow> rpt_rows(const RptPage &pg, double end)
{
    /* a row: the things that start before the first of them ends (VB6 oracle d_rpt2: "right" at 0 and a
     * bigger "Header" at 100 are one row, "center" at 260 the next) */
    std::vector<const RptItem *> all;
    for (const RptItem &it : pg.items) if ((it.kind == 1 && it.h < 1) || (it.kind == 0 && !it.text.isEmpty())) all.push_back(&it);
    std::stable_sort(all.begin(), all.end(), [](const RptItem *a, const RptItem *b) { return a->y < b->y - 0.5; });
    std::vector<std::vector<const RptItem *>> groups;
    std::vector<double> gy;
    for (size_t i = 0; i < all.size();) {
        const RptItem *f = all[i];
        double lim = f->y + std::max(f->h, 1.0);
        std::vector<const RptItem *> g;
        for (; i < all.size() && all[i]->y < lim - 0.5; i++) g.push_back(all[i]);
        if (g.empty()) { g.push_back(all[i]); i++; }
        groups.push_back(g); gy.push_back(f->y);
    }
    std::vector<RRow> rows;
    for (size_t gi = 0; gi < groups.size(); gi++) {
        double y = gy[gi], next = gi + 1 < groups.size() ? gy[gi + 1] : end;
        std::vector<const RptItem *> lines, texts;
        for (const RptItem *it : groups[gi]) (it->kind == 1 ? lines : texts).push_back(it);
        auto under = [&](const RptItem *t) {
            for (const RptItem *l : lines) if (t->x < l->x + l->w - 0.5 && t->x + t->w > l->x + 0.5) return true;
            return false;
        };
        if (lines.empty()) { rows.push_back({ y, next - y, texts, {} }); continue; }
        RRow a { y, 0, lines, {} }, b { y, next - y, {}, {} };
        for (const RptItem *t : texts) (under(t) ? b.items : a.items).push_back(t);
        if (b.items.empty()) { a.h = next - y; rows.push_back(a); continue; }
        for (const RptItem *t : a.items) if (t->kind == 0) a.spanned.push_back(t);
        rows.push_back(a); rows.push_back(b);
    }
    for (RRow &r : rows) std::sort(r.items.begin(), r.items.end(), [](const RptItem *x, const RptItem *y) { return x->x < y->x; });
    return rows;
}
/* a page as an HTML table, as VB6 writes it */
static QString rpt_html(Report *r, const RptPage &pg, double ph)
{
    double total = r->lm + (r->width > 0 ? r->width : 9360);
    std::vector<double> xs { 0, r->lm };
    for (const RptItem &it : pg.items) if (it.kind == 1 || !it.text.isEmpty()) xs.push_back(it.x);
    std::sort(xs.begin(), xs.end());
    xs.erase(std::unique(xs.begin(), xs.end(), [](double a, double b) { return std::fabs(a - b) < 1; }), xs.end());
    if (xs.back() < total) xs.push_back(total);
    QString out = "<TABLE WIDTH=100%>\r\n<TR><TD WIDTH=0></TD>";
    for (size_t i = 1; i + 1 < xs.size(); i++)
        out += QString("<TH WIDTH=%1%></TH>").arg((int)std::lround((xs[i + 1] - xs[i]) * 100 / total));
    out += "</TR>\r\n";
    auto col_of = [&](double x) { int k = 0; for (size_t i = 1; i < xs.size(); i++) if (xs[i] <= x + 0.5) k = (int)i; return k; };
    std::vector<RRow> rows = rpt_rows(pg, ph - r->bm);
    double top = rows.empty() ? ph : rows[0].y;
    if (top > 0) out += QString("  <TR><TD HEIGHT=%1></TD></TR>\r\n").arg((int)std::lround(top / 15));
    double last = top;
    for (const RRow &row : rows) {
        out += QString("  <TR><TD HEIGHT=%1></TD>").arg((int)std::lround(row.h / 15));
        int at = 1;
        for (const RptItem *it : row.items) {
            int c0 = col_of(it->x), c1 = col_of(it->x + std::max(it->w, 1.0) - 1);
            if (c0 < at) continue;                  /* two things in one cell: the first */
            for (; at < c0; at++) out += "<TD></TD>";
            int span = std::max(1, c1 - c0 + 1);
            QString cell = span > 1 ? QString("<TD COLSPAN=%1").arg(span) : QString("<TD");
            if (std::find(row.spanned.begin(), row.spanned.end(), it) != row.spanned.end()) cell += " ROWSPAN=2";
            if (it->kind == 1) cell += " VALIGN=TOP><HR></TD>";
            else {
                QString t = html_escape(it->text);
                if (it->font.bold) t = "<B>" + t + "</B>";
                if (it->font.italic) t = "<I>" + t + "</I>";
                if (it->font.underline) t = "<U>" + t + "</U>";
                if (it->align == 1) t = "<DIV ALIGN=RIGHT>" + t + "</DIV>";
                else if (it->align == 2) t = "<CENTER>" + t + "</CENTER>";
                cell += QString(" NOWRAP VALIGN=TOP><FONT FACE=\"%1\" SIZE=%2>%3</FONT></TD>").arg(it->font.name).arg(html_size(it->font.size)).arg(t);
            }
            out += cell;
            at = c0 + span;
        }
        out += "</TR>\r\n";
        last = row.y + row.h;
    }
    if (last < ph) out += QString("  <TR><TD HEIGHT=%1></TD></TR>\r\n").arg((int)std::lround((ph - last) / 15));
    out += "</TABLE>\r\n";
    return out;
}
/* a page as lines of text: a character is 75 twips wide, a row one line or more of 230.4 twips
 * (VB6 oracles d_rpt1, d_rpt2) */
static QString rpt_textpage(Report *r, const RptPage &pg, double ph)
{
    const double cw = 75, lh = 230.4;
    int c0 = (int)std::floor(r->lm / cw);
    std::vector<QString> lines;
    auto put = [&](int line, int col, const QString &s) {
        while ((int)lines.size() <= line) lines.push_back(QString());
        QString &l = lines[line];
        if (col < 0) col = 0;
        if (l.size() < col) l += QString(col - l.size(), ' ');
        if (l.size() > col) { if (l.size() > col + s.size()) l.replace(col, s.size(), s); else l = l.left(col) + s; }
        else l += s;
    };
    std::vector<RRow> rows = rpt_rows(pg, ph - r->bm);
    int line = rows.empty() ? 0 : std::max(0, (int)std::lround(rows[0].y / lh));
    for (size_t ri = 0; ri < rows.size(); ri++) {
        const RRow &row = rows[ri];
        if (ri) line += std::max(1, (int)std::lround(rows[ri - 1].h / lh));
        for (const RptItem *it : row.items) {
            double x = it->x - r->lm;
            if (it->kind == 1) { put(line, c0 + (int)std::lround(x / cw), QString((int)std::floor(it->w / cw), '-')); continue; }
            int n = it->text.size(), col;
            if (it->align == 1) col = c0 + (int)std::floor((x + it->w) / cw) - n;
            else if (it->align == 2) col = c0 + (int)std::lround((x + it->w / 2) / cw - n / 2.0);
            else col = c0 + (int)std::lround(x / cw);
            put(line, col, it->text);
        }
    }
    if (!rows.empty()) line += std::max(1, (int)std::lround(rows.back().h / lh));
    int total = line + std::max(0, (int)std::lround((ph - (rows.empty() ? 0 : rows.back().y + rows.back().h)) / lh));
    while ((int)lines.size() < total) lines.push_back(QString());
    QString out;
    for (const QString &l : lines) { QString t = l; while (t.endsWith(' ')) t.chop(1); out += t + "\r\n"; }
    return out;
}
static RptFormat *rpt_format(Report *r, vbvar key)
{
    vbvar k = vb_v_letval(key);
    if (k.vt == VT_BSTR) {
        QString s = qs(k.u.s);
        for (RptFormat *f : r->formats) if (!f->key.compare(s, Qt::CaseInsensitive)) return f;
        return nullptr;
    }
    int i = vb_v2l(k);
    return i >= 1 && i <= (int)r->formats.size() ? r->formats[i - 1] : nullptr;
}
static void rpt_export(Report *r, vb_i32 argc, vbvar *argv)
{
    auto miss = [&](int i) { return i >= argc || vb_ismissing(argv[i]); };
    RptFormat *f = miss(0) ? r->formats[0] : rpt_format(r, argv[0]);
    if (!f) { rpt_raise(8501, "ExportFormats", QString("Key not found: '%1'").arg(qs(vb_v2s(argv[0])))); return; }
    QString file = miss(1) ? QString() : qs(vb_v2s(argv[1]));
    bool overwrite = miss(2) ? true : vb_v_truth(argv[2]);
    bool dialog = miss(3) ? false : vb_v_truth(argv[3]);
    int range = miss(4) ? 0 : vb_v2l(argv[4]);
    int from = miss(5) ? 1 : vb_v2l(argv[5]), to = miss(6) ? 1 << 30 : vb_v2l(argv[6]);
    if (dialog && !monkey_on) {
        QString filter = QString("%1 (%2)").arg(f->filestring, f->filter);
        QString got = QFileDialog::getSaveFileName(nullptr, "Export", file, f->filestring);
        if (got.isEmpty()) return;
        file = got;
    }
    if (file.isEmpty()) { rpt_raise(8519, "DataReport", "Invalid file name."); return; }
    QString path = file;
#ifndef _WIN32
    path.replace('\\', '/');
#endif
    if (QFileInfo(path).suffix().isEmpty()) {           /* the extension of the format */
        QString ext = f->filter.section(';', 0, 0).trimmed();
        if (ext.startsWith("*.")) path += ext.mid(1);
    }
    if (!overwrite && QFileInfo::exists(path)) { rpt_raise(8512, "DataReport", QString("File already exists: '%1'").arg(file)); return; }
    if (!rpt_layout(r)) return;
    if (range != 1) { from = 1; to = (int)r->pages.size(); }
    double ph = r->orientation == 2 ? 12240 : 15840;
    QString body;
    for (int p = std::max(1, from); p <= std::min(to, (int)r->pages.size()); p++)
        body += (f->type == 0 || f->type == 3) ? rpt_html(r, r->pages[p - 1], ph) : rpt_textpage(r, r->pages[p - 1], ph);
    QString text = f->tmpl;
    text.replace("<!--MSDBRPT_Template_Title-->", f->type == 0 || f->type == 3 ? html_escape(r->title) : r->title);
    text.replace("<!--MSDBRPT_Template_Body-->", body);
    QFile out(path);
    if (!out.open(QIODevice::WriteOnly | QIODevice::Truncate)) { rpt_raise(8514, "DataReport", "Path not found."); return; }
    QByteArray data;
    if (f->type == 2) { data = QByteArray("\xFF\xFE", 2); data.append(reinterpret_cast<const char *>(text.utf16()), text.size() * 2); }
    else if (f->type == 3) data = text.toUtf8();
    else { vb_i32 m2 = vb_mark(); int n = 0; char *a = vb_stoansi(vs(text), &n); data = QByteArray(a, n); vb_release(m2); }   /* the ANSI code page */
    out.write(data);
}

/* ---------------------------------------------------------------- drawing: the preview and the PDF */
static void rpt_draw(QPainter &p, const RptPage &pg, double scale)
{
    p.setRenderHint(QPainter::Antialiasing);
    for (const RptItem &it : pg.items) {
        QRectF rc(it.x * scale, it.y * scale, it.w * scale, it.h * scale);
        QColor fore(it.fore & 0xFF, (it.fore >> 8) & 0xFF, (it.fore >> 16) & 0xFF);
        QColor back(it.back & 0xFF, (it.back >> 8) & 0xFF, (it.back >> 16) & 0xFF);
        QColor border(it.border & 0xFF, (it.border >> 8) & 0xFF, (it.border >> 16) & 0xFF);
        if (it.kind == 1) {
            p.setPen(QPen(border, std::max(1.0, 15 * scale)));
            if (it.slant) p.drawLine(rc.topRight(), rc.bottomLeft()); else p.drawLine(rc.topLeft(), rc.bottomRight());
            continue;
        }
        if (it.kind == 2) {
            p.setPen(it.borderstyle ? QPen(border, std::max(1.0, 15 * scale)) : Qt::NoPen);
            p.setBrush(it.backstyle ? QBrush(back) : Qt::NoBrush);
            QRectF s = rc;
            if (it.shape == 1 || it.shape == 3 || it.shape == 5) { double d = std::min(rc.width(), rc.height()); s = QRectF(rc.center().x() - d / 2, rc.center().y() - d / 2, d, d); }
            if (it.shape == 2 || it.shape == 3) p.drawEllipse(s);
            else if (it.shape == 4 || it.shape == 5) p.drawRoundedRect(s, s.width() / 8, s.height() / 8);
            else p.drawRect(s);
            p.setBrush(Qt::NoBrush);
            continue;
        }
        if (it.kind == 3) {                     /* RptImage: clip, stretch or zoom, at its PictureAlignment */
            if (it.backstyle) p.fillRect(rc, back);
            if (!it.img.isNull()) {
                double iw = it.img.width() * 15.0 * scale, ih = it.img.height() * 15.0 * scale;   /* 96 dpi: 15 twips a pixel */
                if (it.sizemode == 1) p.drawImage(rc, it.img);
                else {
                    if (it.sizemode == 2) { double k = std::min(rc.width() / iw, rc.height() / ih); iw *= k; ih *= k; }
                    int a = it.picalign;
                    double x = (a == 0 || a == 6 || a == 7) ? rc.left() : (a == 2 || a == 3 || a == 4) ? rc.right() - iw : rc.center().x() - iw / 2;
                    double y = (a == 0 || a == 1 || a == 2) ? rc.top() : (a == 4 || a == 5 || a == 6) ? rc.bottom() - ih : rc.center().y() - ih / 2;
                    p.save();
                    p.setClipRect(rc, Qt::IntersectClip);
                    p.drawImage(QRectF(x, y, iw, ih), it.img);
                    p.restore();
                }
            }
            if (it.borderstyle) { p.setPen(QPen(border, 1)); p.setBrush(Qt::NoBrush); p.drawRect(rc); }
            continue;
        }
        if (it.kind == 0) {
            if (it.backstyle) p.fillRect(rc, back);
            QFont f(it.font.name);
            f.setPointSizeF(std::max(1.0, it.font.size * scale * 20 / 1.0 / 20));
            f.setPixelSize(std::max(1, (int)std::lround(it.font.size * 20 * scale)));
            f.setBold(it.font.bold); f.setItalic(it.font.italic); f.setUnderline(it.font.underline); f.setStrikeOut(it.font.strike);
            p.setFont(f);
            p.setPen(fore);
            int fl = Qt::AlignTop | Qt::TextSingleLine | (it.align == 1 ? Qt::AlignRight : it.align == 2 ? Qt::AlignHCenter : Qt::AlignLeft);
            p.drawText(rc, fl, it.text);
            if (it.borderstyle) { p.setPen(QPen(border, 1)); p.drawRect(rc); }
        }
    }
}
static void rpt_print(Report *r)
{
    if (!rpt_layout(r)) return;
    QString file;
    const char *e = getenv("VBRT_PRINTER_FILE");
    vb_i32 m = vb_mark();
    file = e && *e ? QString::fromUtf8(e) : qs(vb_app_path()) + "/" + qs(vb_app_exename()) + "-report.pdf";
    vb_release(m);
    QPdfWriter pdf(file);
    pdf.setResolution(1440);
    pdf.setPageSize(QPageSize(QPageSize::Letter));
    pdf.setPageOrientation(r->orientation == 2 ? QPageLayout::Landscape : QPageLayout::Portrait);
    pdf.setPageMargins(QMarginsF(0, 0, 0, 0));
    QPainter p(&pdf);
    for (size_t i = 0; i < r->pages.size(); i++) {
        if (i) pdf.newPage();
        rpt_draw(p, r->pages[i], 1.0);
    }
}

/* the preview window */
class ReportView : public QWidget {
public:
    Report *r;
    explicit ReportView(Report *rep) : r(rep) { setAutoFillBackground(true); }
    QSize sizeHint() const override { return QSize(pagew(), pageh()); }
    int pagew() const { return (int)std::lround((r->orientation == 2 ? 15840 : 12240) / 15.0 * r->zoom / 100); }
    int pageh() const { return (int)std::lround((r->orientation == 2 ? 12240 : 15840) / 15.0 * r->zoom / 100); }
protected:
    void paintEvent(QPaintEvent *) override
    {
        QPainter p(this);
        p.fillRect(rect(), QColor(128, 128, 128));
        QRect pg(0, 0, pagew(), pageh());
        p.fillRect(pg, Qt::white);
        if (r->curpage >= 0 && r->curpage < (int)r->pages.size()) rpt_draw(p, r->pages[r->curpage], r->zoom / 100.0 / 15.0);
    }
};
static void rpt_event(Report *r, const char *ev, int n, vbvar *a)
{
    if (!r->owner) return;
    vbvar args[5] = { vb_v_s(vb_sfromutf8(ev)), n > 0 ? a[0] : vb_v_empty(), n > 1 ? a[1] : vb_v_empty(),
                      n > 2 ? a[2] : vb_v_empty(), n > 3 ? a[3] : vb_v_empty() };
    vb_i32 save = vb_errpend;
    vb_late_o(r->owner, "lb_event", VBM_METHOD, 5, args);
    if (vb_errpend && !save) vb_unhandled();
    for (int i = 0; i < n; i++) a[i] = args[i + 1];
}
class ReportWin : public QWidget {
public:
    Report *r;
    QScrollArea *area; ReportView *view; QLabel *pages;
    explicit ReportWin(Report *rep) : QWidget(nullptr, Qt::Window), r(rep)
    {
        setAttribute(Qt::WA_DeleteOnClose, false);
        auto *lay = new QVBoxLayout(this);
        lay->setContentsMargins(0, 0, 0, 0); lay->setSpacing(0);
        auto *tb = new QToolBar(this);
        tb->addAction("Print", [this] { rpt_print(r); });
        tb->addAction("Export", [this] { vbvar a[4] = { vb_v_l(1), vb_v_s(vb_sfromutf8("")), vb_v_bool(-1), vb_v_bool(-1) }; rpt_export(r, 4, a); if (vb_errpend) vb_unhandled(); });
        auto *zoom = new QComboBox(tb);
        for (int z : { 50, 75, 100, 150, 200 }) zoom->addItem(QString("%1%").arg(z), z);
        zoom->setCurrentIndex(2);
        tb->addWidget(zoom);
        tb->addAction("|<", [this] { go(0); });
        tb->addAction("<", [this] { go(r->curpage - 1); });
        pages = new QLabel(tb);
        tb->addWidget(pages);
        tb->addAction(">", [this] { go(r->curpage + 1); });
        tb->addAction(">|", [this] { go((int)r->pages.size() - 1); });
        lay->addWidget(tb);
        area = new QScrollArea(this);
        area->setBackgroundRole(QPalette::Dark);
        view = new ReportView(r);
        area->setWidget(view);
        lay->addWidget(area);
        QObject::connect(zoom, QOverload<int>::of(&QComboBox::currentIndexChanged), [this, zoom](int) {
            r->zoom = zoom->currentData().toInt(); view->resize(view->sizeHint()); view->update(); });
    }
    void go(int p)
    {
        if (r->pages.empty()) return;
        r->curpage = std::max(0, std::min(p, (int)r->pages.size() - 1));
        pages->setText(QString(" %1 / %2 ").arg(r->curpage + 1).arg(r->pages.size()));
        view->resize(view->sizeHint());
        view->update();
    }
protected:
    void closeEvent(QCloseEvent *e) override
    {
        vbvar a[2] = { vb_v_i(0), vb_v_i(0) };  /* Cancel, vbFormControlMenu */
        rpt_event(r, "queryclose", 2, a);
        if (vb_v2l(a[0])) { e->ignore(); return; }
        e->accept();
    }
    void showEvent(QShowEvent *) override { rpt_event(r, "activate", 0, nullptr); }
    void resizeEvent(QResizeEvent *) override { rpt_event(r, "resize", 0, nullptr); }
};
static void rpt_show(Report *r, bool modal)
{
    if (!rpt_layout(r)) return;
    if (!r->win) {
        r->win = new ReportWin(r);
        double tpp = 15;
        int w = r->ww > 0 ? (int)(r->ww / tpp) : 800, h = r->wh > 0 ? (int)(r->wh / tpp) : 600;
        r->win->resize(w, h);
        if (r->wl >= 0 && r->wt >= 0) r->win->move((int)(r->wl / tpp), (int)(r->wt / tpp));
    }
    r->win->setWindowTitle(r->caption);
    ((ReportWin *)r->win)->go(r->curpage);
    if (monkey_on) { monkey_log("DataReport " + r->caption + QString(": %1 pages").arg(r->pages.size())); }
    if (const char *dir = getenv("VBRT_REPORT_SHOTS")) {     /* the test robot: pictures of the first pages */
        for (int p = 0; p < (int)r->pages.size() && p < 3; p++) {
            QImage img(816, 1056, QImage::Format_RGB32);
            img.fill(Qt::white);
            QPainter pt(&img);
            rpt_draw(pt, r->pages[p], 1 / 15.0);
            pt.end();
            img.save(QString("%1/report-%2-page%3.png").arg(QString::fromUtf8(dir), r->caption).arg(p + 1));
        }
    }
    if (r->windowstate == 2) r->win->showMaximized(); else if (r->windowstate == 1) r->win->showMinimized(); else r->win->show();
    if (modal) {
        QEventLoop loop;
        while (r->win && r->win->isVisible()) { QApplication::processEvents(QEventLoop::WaitForMoreEvents); }
    }
}

/* ---------------------------------------------------------------- the objects */
static vbvar rcolor(long c) { return vb_v_l((vb_i32)c); }
static vbobj ctl_obj(RptCtl *c)
{
    if (!c->obj) {
        const vb_class *k = c->type == "RptTextBox" ? &CLS_RTEXTBOX : c->type == "RptLine" ? &CLS_RLINE :
                            c->type == "RptShape" ? &CLS_RSHAPE : c->type == "RptImage" ? &CLS_RIMAGE :
                            c->type == "RptFunction" ? &CLS_RFUNCTION : &CLS_RLABEL;
        c->obj = vb_onew(k);
        c->obj->ext = c;
        vb_oaddref(c->obj);
    }
    return c->obj;
}
static vbvar rpt_obj(vbobj o) { return vb_vtmp(vb_v_obj(o)); }
static vbvar rfont_obj(RFont &f)
{
    vbobj o = vb_onew(vbf_cls_stdfont());
    vb_i32 m = vb_mark();
    vbvar v;
    v = vb_v_s(vs(f.name)); vb_late_o(o, "name", VBM_LET, 1, &v);
    v = vb_v_d(f.size); vb_late_o(o, "size", VBM_LET, 1, &v);
    v = vb_v_bool(f.bold ? -1 : 0); vb_late_o(o, "bold", VBM_LET, 1, &v);
    v = vb_v_bool(f.italic ? -1 : 0); vb_late_o(o, "italic", VBM_LET, 1, &v);
    v = vb_v_bool(f.underline ? -1 : 0); vb_late_o(o, "underline", VBM_LET, 1, &v);
    vb_release(m);
    return rpt_obj(o);
}
static void rfont_from(RFont &f, vbvar v)
{
    vbobj o = v.vt == VT_DISPATCH ? v.u.o : nullptr;
    if (!o) return;
    f.name = qs(vb_v2s(vb_late_o(o, "name", VBM_GET, 0, nullptr)));
    f.size = vb_v2d(vb_late_o(o, "size", VBM_GET, 0, nullptr));
    f.bold = vb_v_truth(vb_late_o(o, "bold", VBM_GET, 0, nullptr));
    f.italic = vb_v_truth(vb_late_o(o, "italic", VBM_GET, 0, nullptr));
    f.underline = vb_v_truth(vb_late_o(o, "underline", VBM_GET, 0, nullptr));
}
static vbvar rptctl_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    RptCtl *c = (RptCtl *)vb_oid(o)->ext;
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
#define NUMP(key, field) if (n == key) { if (set) { c->field = vb_v2d(val); return vb_v_empty(); } return vb_v_f((float)c->field); }
#define INTP(key, field) if (n == key) { if (set) { c->field = vb_v2l(val); return vb_v_empty(); } return vb_v_l(c->field); }
#define BOOLP(key, field) if (n == key) { if (set) { c->field = vb_v_truth(val); return vb_v_empty(); } return vb_v_bool(c->field ? -1 : 0); }
#define STRP(key, field) if (n == key) { if (set) { c->field = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(c->field); }
    if (n == "name") { if (set) { rpt_raise(8579, "DataReport", "Property is read-only at run time."); return vb_v_empty(); } return vstr(c->name); }
    NUMP("left", left) NUMP("top", top) NUMP("width", width) NUMP("height", height)
    BOOLP("visible", visible)
    if (n == "bordercolor") { if (set) { c->bordercolor = vb_v2l(val); return vb_v_empty(); } return rcolor(c->bordercolor); }
    INTP("borderstyle", borderstyle)
    if (c->type == "RptLine") { INTP("lineslant", lineslant) }
    if (c->type == "RptShape" || c->type == "RptImage" || c->type == "RptLabel" || c->type == "RptTextBox" || c->type == "RptFunction") {
        if (n == "backcolor") { if (set) { c->backcolor = vb_v2l(val); return vb_v_empty(); } return rcolor(c->backcolor); }
        INTP("backstyle", backstyle)
    }
    if (c->type == "RptShape") { INTP("shape", shape) }
    if (c->type == "RptImage") {
        INTP("sizemode", sizemode) INTP("picturealignment", picalign)
        if (n == "picture") {
            if (set) {
                const QImage *im = image_of_value(val);
                c->img = im ? *im : QImage();
                return vb_v_empty();
            }
            if (c->img.isNull()) return rpt_obj(nullptr);
            return rpt_obj(new_pic(c->img, 1));
        }
    }
    if (c->type == "RptLabel" || c->type == "RptTextBox" || c->type == "RptFunction") {
        if (n == "forecolor") { if (set) { c->forecolor = vb_v2l(val); return vb_v_empty(); } return rcolor(c->forecolor); }
        INTP("alignment", alignment) BOOLP("cangrow", cangrow) BOOLP("righttoleft", rtl)
        if (n == "font") { if (set) { rfont_from(c->font, val); return vb_v_empty(); } return rfont_obj(c->font); }
    }
    if (c->type == "RptLabel") { STRP("caption", caption) }
    if (c->type == "RptTextBox" || c->type == "RptFunction") {
        STRP("datafield", datafield) STRP("datamember", datamember)
        if (n == "dataformat") return rpt_obj(nullptr);
    }
    if (c->type == "RptFunction") { INTP("functiontype", functype) }
#undef NUMP
#undef INTP
#undef BOOLP
#undef STRP
    *found = 0;
    return vb_v_empty();
}
static int find_in(vbvar key, int n, const std::function<QString(int)> &nameof)
{
    vbvar k = vb_v_letval(key);
    if (k.vt == VT_BSTR) { QString s = qs(k.u.s); for (int i = 0; i < n; i++) if (!nameof(i).compare(s, Qt::CaseInsensitive)) return i; return -1; }
    int i = vb_v2l(k) - 1;
    return i >= 0 && i < n ? i : -1;
}
static vbvar rctls_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    RptSec *s = (RptSec *)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    if (n == "count") return vb_v_l((vb_i32)s->ctls.size());
    if (n == "item" || n.empty()) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = find_in(argv[0], (int)s->ctls.size(), [s](int k) { return s->ctls[k]->name; });
        if (i < 0) { rpt_raise(8574, "DataReport", "Object not found in collection."); return vb_v_empty(); }
        return rpt_obj(ctl_obj(s->ctls[i]));
    }
    *found = 0;
    return vb_v_empty();
}
static vbvar rsec_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    RptSec *s = (RptSec *)vb_oid(o)->ext;
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (n == "name") { if (set) { rpt_raise(8579, "DataReport", "Property is read-only at run time."); return vb_v_empty(); } return vstr(s->name); }
    if (n == "height") { if (set) { s->height = vb_v2d(val); return vb_v_empty(); } return vb_v_f((float)s->height); }
    if (n == "keeptogether") { if (set) { s->keep = vb_v_truth(val); return vb_v_empty(); } return vb_v_bool(s->keep ? -1 : 0); }
    if (n == "forcepagebreak") { if (set) { s->forcebreak = vb_v2l(val); return vb_v_empty(); } return vb_v_l(s->forcebreak); }
    if (n == "visible") { if (set) { s->visible = vb_v_truth(val); return vb_v_empty(); } return vb_v_bool(s->visible ? -1 : 0); }
    if (n == "controls") {
        if (!s->ctlsobj) { s->ctlsobj = vb_onew(&CLS_RCONTROLS); s->ctlsobj->ext = s; vb_oaddref(s->ctlsobj); }
        if (argc && !set) { int f2; return rctls_member(s->ctlsobj, "item", VBM_GET, argc, argv, &f2); }
        return rpt_obj(s->ctlsobj);
    }
    *found = 0;
    return vb_v_empty();
}
static vbvar rsecs_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    Report *r = (Report *)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    if (n == "count") return vb_v_l((vb_i32)r->secs.size());
    if (n == "item" || n.empty()) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = find_in(argv[0], (int)r->secs.size(), [r](int k) { return r->secs[k]->name; });
        if (i < 0) { rpt_raise(8574, "DataReport", "Object not found in collection."); return vb_v_empty(); }
        RptSec *s = r->secs[i];
        if (!s->obj) { s->obj = vb_onew(&CLS_RSECTION); s->obj->ext = s; vb_oaddref(s->obj); }
        return rpt_obj(s->obj);
    }
    *found = 0;
    return vb_v_empty();
}
static vbvar rfmt_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    RptFormat *f = (RptFormat *)vb_oid(o)->ext;
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (n == "key") { if (set) { f->key = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(f->key); }
    if (n == "template") { if (set) { f->tmpl = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(f->tmpl); }
    if (n == "formattype") { if (set) { f->type = vb_v2l(val); return vb_v_empty(); } return vb_v_l(f->type); }
    if (n == "fileformatstring") { if (set) { f->filestring = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(f->filestring); }
    if (n == "filefilter") { if (set) { f->filter = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(f->filter); }
    *found = 0;
    return vb_v_empty();
}
static vbobj fmt_obj(RptFormat *f)
{
    if (!f->obj) { f->obj = vb_onew(&CLS_RFORMAT); f->obj->ext = f; vb_oaddref(f->obj); }
    return f->obj;
}
static vbvar rfmts_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    Report *r = (Report *)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    auto miss = [&](int i) { return i >= argc || vb_ismissing(argv[i]); };
    if (n == "count") return vb_v_l((vb_i32)r->formats.size());
    if (n == "item" || n.empty()) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        RptFormat *f = rpt_format(r, argv[0]);
        if (!f) { rpt_raise(8501, "ExportFormats", QString("Key not found: '%1'").arg(qs(vb_v2s(argv[0])))); return vb_v_empty(); }
        return rpt_obj(fmt_obj(f));
    }
    if (n == "add" || n == "insert") {
        int off = n == "insert" ? 1 : 0;
        QString key = miss(off) ? QString() : qs(vb_v2s(argv[off]));
        for (RptFormat *g : r->formats)
            if (!key.isEmpty() && !g->key.compare(key, Qt::CaseInsensitive)) { rpt_raise(8500, "ExportFormats", QString("Duplicate key value '%1'").arg(key)); return vb_v_empty(); }
        RptFormat *f = new RptFormat;
        f->key = key;
        f->type = miss(off + 1) ? 0 : vb_v2l(argv[off + 1]);
        f->filestring = miss(off + 2) ? QString() : qs(vb_v2s(argv[off + 2]));
        f->filter = miss(off + 3) ? QString() : qs(vb_v2s(argv[off + 3]));
        f->tmpl = miss(off + 4) ? QString() : qs(vb_v2s(argv[off + 4]));
        if (f->tmpl.isEmpty()) {                /* the template of its type */
            for (RptFormat *g : r->formats) if (g->key.startsWith("key_def_") && g->type == f->type) { f->tmpl = g->tmpl; break; }
        }
        int at = (int)r->formats.size();
        if (off) { at = vb_v2l(argv[0]) - 1; if (at < 0 || at > (int)r->formats.size()) { vb_raise(9); delete f; return vb_v_empty(); } }
        r->formats.insert(r->formats.begin() + at, f);
        return rpt_obj(fmt_obj(f));
    }
    if (n == "remove") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        RptFormat *f = rpt_format(r, argv[0]);
        if (!f) { rpt_raise(8501, "ExportFormats", "Key not found"); return vb_v_empty(); }
        r->formats.erase(std::find(r->formats.begin(), r->formats.end(), f));
        return vb_v_empty();
    }
    if (n == "clear") { r->formats.clear(); return vb_v_empty(); }
    *found = 0;
    return vb_v_empty();
}
static vbvar rep_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    Report *r = REP(o);
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    auto miss = [&](int i) { return i >= argc || vb_ismissing(argv[i]); };
    *found = 1;
    if (n == "caption") { if (set) { r->caption = qs(vb_v2s(val)); if (r->win) r->win->setWindowTitle(r->caption); return vb_v_empty(); } return vstr(r->caption); }
    if (n == "title") { if (set) { r->title = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(r->title); }
    if (n == "datamember") { if (set) { r->datamember = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(r->datamember); }
    if (n == "datasource") {
        if (set) {
            vbobj ds = val.vt == VT_DISPATCH ? val.u.o : nullptr;
            if (ds) vb_oaddref(ds);
            if (r->datasource) vb_orelease(r->datasource);
            r->datasource = ds;
            return vb_v_empty();
        }
        return rpt_obj(r->datasource);
    }
#define TW(key, field) if (n == key) { if (set) { r->field = vb_v2d(val); return vb_v_empty(); } return vb_v_f((float)r->field); }
    TW("leftmargin", lm) TW("rightmargin", rm) TW("topmargin", tm) TW("bottommargin", bm) TW("reportwidth", width)
    TW("gridx", gridx) TW("gridy", gridy)
#undef TW
    if (n == "orientation") { if (set) { r->orientation = vb_v2l(val); return vb_v_empty(); } return vb_v_l(r->orientation); }
    if (n == "font") { if (set) { rfont_from(r->font, val); return vb_v_empty(); } return rfont_obj(r->font); }
    if (n == "asynccount") return vb_v_l(0);    /* the work is done when the call returns */
    if (n == "sections") {
        if (!r->secsobj) { r->secsobj = vb_onew(&CLS_RSECTIONS); r->secsobj->ext = r; vb_oaddref(r->secsobj); }
        if (argc && !set) { int f2; return rsecs_member(r->secsobj, "item", VBM_GET, argc, argv, &f2); }
        return rpt_obj(r->secsobj);
    }
    if (n == "exportformats") {
        if (!r->fmtsobj) { r->fmtsobj = vb_onew(&CLS_RFORMATS); r->fmtsobj->ext = r; vb_oaddref(r->fmtsobj); }
        if (argc && !set) { int f2; return rfmts_member(r->fmtsobj, "item", VBM_GET, argc, argv, &f2); }
        return rpt_obj(r->fmtsobj);
    }
    if (n == "exportreport") { rpt_export(r, argc, argv); return vb_v_l(0); }
    if (n == "printreport") {
        if (!miss(0) && vb_v_truth(argv[0]) && !monkey_on) {}
        rpt_print(r);
        return vb_v_l(0);
    }
    if (n == "refresh") { if (r->win && r->win->isVisible()) { rpt_layout(r); ((ReportWin *)r->win)->go(r->curpage); } return vb_v_empty(); }
    if (n == "show") { rpt_show(r, !miss(0) && vb_v2l(argv[0]) == 1); return vb_v_empty(); }
    if (n == "hide") { if (r->win) r->win->hide(); return vb_v_empty(); }
    if (n == "lb_picture" && argc >= 2) {      /* the Picture of an RptImage, from the .dsx (base64) */
        QString who = qs(vb_v2s(argv[0])).toLower();
        QImage img;
        img.loadFromData(QByteArray::fromBase64(qs(vb_v2s(argv[1])).toLatin1()));
        for (RptSec *s : r->secs) for (RptCtl *c : s->ctls) if (c->name.toLower() == who) c->img = img;
        return vb_v_empty();
    }
    if (n == "lb_close") { if (r->win) { r->win->hide(); r->win->deleteLater(); r->win = nullptr; } return vb_v_empty(); }
    if (n == "visible") {
        if (set) { if (vb_v_truth(val)) rpt_show(r, false); else if (r->win) r->win->hide(); return vb_v_empty(); }
        return vb_v_bool(r->win && r->win->isVisible() ? -1 : 0);
    }
    if (n == "left" || n == "top" || n == "width" || n == "height") {
        double &f = n == "left" ? r->wl : n == "top" ? r->wt : n == "width" ? r->ww : r->wh;
        if (set) {
            f = vb_v2d(val);
            if (r->win) {
                if (n == "left" || n == "top") r->win->move((int)(std::max(0.0, r->wl) / 15), (int)(std::max(0.0, r->wt) / 15));
                else r->win->resize((int)(std::max(100.0, r->ww) / 15), (int)(std::max(100.0, r->wh) / 15));
            }
            return vb_v_empty();
        }
        return vb_v_f((float)(f < 0 ? 0 : f));
    }
    if (n == "windowstate") { if (set) { r->windowstate = vb_v2l(val); return vb_v_empty(); } return vb_v_l(r->windowstate); }
    if (n == "hwnd") return vb_v_l(r->win ? (vb_i32)r->win->winId() : 0);
    if (n == "aboutbox" || n == "cancel" || n == "zorder") return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}
static void rep_clear(vbobj o)
{
    Report *r = (Report *)o->ext;
    if (!r) return;
    if (r->win) { r->win->hide(); r->win->deleteLater(); }
    if (r->datasource) vb_orelease(r->datasource);
    for (RptSec *s : r->secs) {
        for (RptCtl *c : s->ctls) { if (c->obj) { c->obj->ext = nullptr; vb_orelease(c->obj); } delete c; }
        if (s->obj) vb_orelease(s->obj);
        if (s->ctlsobj) vb_orelease(s->ctlsobj);
        delete s;
    }
    for (RptFormat *f : r->formats) { if (f->obj) vb_orelease(f->obj); delete f; }
    if (r->secsobj) vb_orelease(r->secsobj);
    if (r->fmtsobj) vb_orelease(r->fmtsobj);
    delete r;
    o->ext = nullptr;
}


/* ---------------------------------------------------------------- SysInfo (SysInfo.ocx): what the system is
 * (the screen's work area, the power, the version) and its events: DisplayChanged, SettingChanged (the
 * work area: SPI_SETWORKAREA 47), SysColorsChanged, TimeChanged, PowerStatusChanged, PowerResume,
 * DeviceArrival / DeviceRemoveComplete of a volume. No window; the DataRept sample of VB98 sizes its
 * report with it. On Linux the power comes from /sys/class/power_supply (or VBRT_POWER_SUPPLY_DIR) and
 * the volumes are the mounted disks (DeviceName: the mount point, DeviceID 0) */
#include <QElapsedTimer>
#include <time.h>
#ifdef _WIN32
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>
#endif
extern "C" int vbos_drives(char *out, int cap);
struct SysPower { int ac = 255, status = 255, percent = 255; vb_i32 life = -1, full = -1; };
static SysPower sys_power()
{
    SysPower p;
#ifdef _WIN32
    SYSTEM_POWER_STATUS s;
    if (GetSystemPowerStatus(&s)) {
        p.ac = s.ACLineStatus; p.status = s.BatteryFlag; p.percent = s.BatteryLifePercent;
        p.life = (vb_i32)s.BatteryLifeTime; p.full = (vb_i32)s.BatteryFullLifeTime;
    }
#else
    const char *env = getenv("VBRT_POWER_SUPPLY_DIR");
    QDir dir(env && *env ? env : "/sys/class/power_supply");
    auto rd = [](const QString &f) { QFile x(f); return x.open(QIODevice::ReadOnly) ? QString::fromLatin1(x.readAll()).trimmed() : QString(); };
    bool any_mains = false, online = false, bat = false;
    double cap = -1, en = -1, full = -1, pw = -1;
    QString st;
    for (const QString &e : dir.entryList(QDir::Dirs | QDir::NoDotAndDotDot)) {
        QString d = dir.filePath(e) + "/";
        QString type = rd(d + "type");
        if (type == "Mains") { any_mains = true; if (rd(d + "online") == "1") online = true; }
        else if (type == "Battery" && !bat && rd(d + "present") != "0") {
            bat = true;
            bool ok; double v = rd(d + "capacity").toDouble(&ok); if (ok) cap = v;
            st = rd(d + "status");
            v = rd(d + "energy_now").toDouble(&ok); if (!ok) v = rd(d + "charge_now").toDouble(&ok); if (ok) en = v;
            v = rd(d + "energy_full").toDouble(&ok); if (!ok) v = rd(d + "charge_full").toDouble(&ok); if (ok) full = v;
            v = rd(d + "power_now").toDouble(&ok); if (!ok) v = rd(d + "current_now").toDouble(&ok); if (ok) pw = v;
        }
    }
    if (!bat) {                                  /* a desktop: AC power, no battery */
        p.ac = 1; p.status = 128; p.percent = 255;
        return p;
    }
    p.ac = any_mains ? (online ? 1 : 0) : (st == "Discharging" ? 0 : 1);
    if (cap >= 0) p.percent = (int)cap;
    /* BatteryFlag of Windows: 1 high (more than 66%), 2 low (less than 33%), 4 critical (less than 5%),
     * 0 between them, 8 charging */
    p.status = p.percent == 255 ? 255 : p.percent > 66 ? 1 : p.percent < 5 ? 4 : p.percent < 33 ? 2 : 0;
    if (st == "Charging") p.status |= 8;
    if (st == "Discharging" && pw > 0 && en >= 0) p.life = (vb_i32)(en / pw * 3600);
    if (pw > 0 && full > 0 && st == "Discharging") p.full = (vb_i32)(full / pw * 3600);
#endif
    return p;
}
static long long sys_slept_ms()                  /* the time the computer spent suspended */
{
#ifdef _WIN32
    ULONGLONG unbiased = 0;
    QueryUnbiasedInterruptTime(&unbiased);
    return (long long)GetTickCount64() - (long long)(unbiased / 10000);
#else
    struct timespec b, m;
    clock_gettime(CLOCK_BOOTTIME, &b); clock_gettime(CLOCK_MONOTONIC, &m);
    return (b.tv_sec - m.tv_sec) * 1000LL + (b.tv_nsec - m.tv_nsec) / 1000000;
#endif
}
static QStringList sys_volumes()
{
    char buf[16384];
    int n = vbos_drives(buf, sizeof buf);
    QStringList out;
    for (char *p = buf; p < buf + n && *p; p += strlen(p) + 1) out << QString::fromUtf8(p);
    return out;
}
#ifndef _WIN32
#include <dlfcn.h>
#include <unistd.h>
/* systemd-logind on the system bus, through libdbus-1 (loaded when the computer has it): the signal
 * PrepareForSleep(true) before the computer sleeps, (false) after it woke up; a "delay" inhibitor lock
 * keeps the sleep waiting (InhibitDelayMaxSec, 5 seconds) while the program runs PowerQuerySuspend
 * and PowerSuspend. DBUS_SYSTEM_BUS_ADDRESS can name another bus (the tests) */
struct SysSleep {
    struct Err { const char *name, *message; unsigned bits; void *pad; void *pad2; };
    void *lib = nullptr, *conn = nullptr;
    int fd = -1;
    void (*err_init)(Err *) = nullptr; void (*err_free)(Err *) = nullptr; unsigned (*err_set)(const Err *) = nullptr;
    void *(*bus_get_private)(int, Err *) = nullptr; void (*set_exit)(void *, unsigned) = nullptr;
    void (*add_match)(void *, const char *, Err *) = nullptr; unsigned (*read_write)(void *, int) = nullptr;
    void *(*pop)(void *) = nullptr; unsigned (*is_signal)(void *, const char *, const char *) = nullptr;
    unsigned (*get_args)(void *, Err *, int, ...) = nullptr; void (*unref)(void *) = nullptr;
    void *(*new_call)(const char *, const char *, const char *, const char *) = nullptr;
    unsigned (*append_args)(void *, int, ...) = nullptr;
    void *(*send_block)(void *, void *, int, Err *) = nullptr;
    void (*conn_close)(void *) = nullptr; void (*conn_unref)(void *) = nullptr;
    ~SysSleep() { release(); if (conn) { conn_close(conn); conn_unref(conn); } }
    template <class F> bool sym(F &f, const char *n) { f = (F)dlsym(lib, n); return f != nullptr; }
    bool open()
    {
        if (getenv("VBRT_NO_LOGIND")) return false;
        lib = dlopen("libdbus-1.so.3", RTLD_NOW | RTLD_LOCAL);
        if (!lib) return false;
        if (!(sym(err_init, "dbus_error_init") && sym(err_free, "dbus_error_free") && sym(err_set, "dbus_error_is_set") &&
              sym(bus_get_private, "dbus_bus_get_private") && sym(set_exit, "dbus_connection_set_exit_on_disconnect") &&
              sym(add_match, "dbus_bus_add_match") && sym(read_write, "dbus_connection_read_write") &&
              sym(pop, "dbus_connection_pop_message") && sym(is_signal, "dbus_message_is_signal") &&
              sym(get_args, "dbus_message_get_args") && sym(unref, "dbus_message_unref") &&
              sym(new_call, "dbus_message_new_method_call") && sym(append_args, "dbus_message_append_args") &&
              sym(send_block, "dbus_connection_send_with_reply_and_block") && sym(conn_close, "dbus_connection_close") &&
              sym(conn_unref, "dbus_connection_unref")))
            return false;
        Err e; err_init(&e);
        conn = bus_get_private(1 /* DBUS_BUS_SYSTEM */, &e);
        if (!conn) { err_free(&e); return false; }
        set_exit(conn, 0);
        add_match(conn, "type='signal',interface='org.freedesktop.login1.Manager',member='PrepareForSleep'", &e);
        if (err_set(&e)) { err_free(&e); return false; }
        inhibit();
        return true;
    }
    void inhibit()
    {
        if (fd >= 0 || !conn) return;
        void *m = new_call("org.freedesktop.login1", "/org/freedesktop/login1", "org.freedesktop.login1.Manager", "Inhibit");
        if (!m) return;
        vb_i32 mk = vb_mark();
        QByteArray who = qs(vb_app_exename()).toUtf8();       /* the program, as the inhibitors list it */
        vb_release(mk);
        const char *what = "sleep", *w = who.constData(), *why = "SysInfo: PowerSuspend", *mode = "delay";
        append_args(m, 's', &what, 's', &w, 's', &why, 's', &mode, 0);
        Err e; err_init(&e);
        void *r = send_block(conn, m, 2000, &e);
        unref(m);
        if (!r) { err_free(&e); return; }
        int f = -1;
        if (get_args(r, &e, 'h', &f, 0)) fd = f; else err_free(&e);
        unref(r);
    }
    void release() { if (fd >= 0) { ::close(fd); fd = -1; } }
    /* the next PrepareForSleep: 1 going to sleep, 0 woke up, -1 none */
    int next()
    {
        if (!conn) return -1;
        read_write(conn, 0);
        while (void *m = pop(conn)) {
            int r = -1;
            if (is_signal(m, "org.freedesktop.login1.Manager", "PrepareForSleep")) {
                unsigned b = 0; Err e; err_init(&e);
                if (get_args(m, &e, 'b', &b, 0)) r = b ? 1 : 0; else err_free(&e);
            }
            unref(m);
            if (r >= 0) return r;
        }
        return -1;
    }
};
#else
#include <QAbstractNativeEventFilter>
#endif
struct OcxSysInfo;
struct SysWatch : QObject {
    OcxSysInfo *si;
    explicit SysWatch(OcxSysInfo *s) : si(s) {}
    bool eventFilter(QObject *o, QEvent *e) override;
};
struct OcxSysInfo : Ocx {
    SysWatch *watch = nullptr;
    QTimer *timer = nullptr, *sleeptimer = nullptr;
#ifndef _WIN32
    SysSleep *sleep = nullptr;                  /* logind: PowerQuerySuspend / PowerSuspend / PowerResume */
#else
    QAbstractNativeEventFilter *native = nullptr;    /* WM_POWERBROADCAST */
    unsigned long long last_resume = 0;
#endif
    SysPower power;
    QStringList vols;
    long long wall0 = 0, slept0 = 0;
    QElapsedTimer mono;
    ~OcxSysInfo() override;
    void going_to_sleep()
    {
        vbvar a[1] = { vb_v_bool(0) };
        fire("powerquerysuspend", 1, a);        /* its Cancel cannot stop the sleep (as since Windows Vista) */
        fire("powersuspend");
    }
    void fire(const char *ev, int argc = 0, vbvar *argv = nullptr) { if (c) fire_ctl(c, ev, argc, argv); }
    void device(const char *ev, const QString &vol)
    {
        vb_i32 id = 0;
#ifdef _WIN32
        if (!vol.isEmpty() && vol[0].isLetter()) id = 1 << (vol[0].toUpper().unicode() - 'A');
        QString name;
#else
        QString name = vol;
#endif
        vbvar a[4] = { vb_v_l(2), vb_v_l(id), vb_v_s(vb_sfromutf8(name.toUtf8().constData())), vb_v_l(0) };
        fire(ev, 4, a);
    }
    void screen_signals(QScreen *sc)
    {
        QObject::connect(sc, &QScreen::geometryChanged, watch, [this](const QRect &) { fire("displaychanged"); });
        QObject::connect(sc, &QScreen::availableGeometryChanged, watch, [this](const QRect &) {
            vbvar a[1] = { vb_v_i(47) };                    /* SPI_SETWORKAREA */
            fire("settingchanged", 1, a);
        });
    }
    void start()
    {
        if (timer) return;
        watch = new SysWatch(this);
        qApp->installEventFilter(watch);
        for (QScreen *sc : QGuiApplication::screens()) screen_signals(sc);
        QObject::connect(qApp, &QGuiApplication::screenAdded, watch, [this](QScreen *sc) { screen_signals(sc); fire("displaychanged"); });
        QObject::connect(qApp, &QGuiApplication::screenRemoved, watch, [this](QScreen *) { fire("displaychanged"); });
        power = sys_power();
        vols = sys_volumes();
        wall0 = QDateTime::currentMSecsSinceEpoch(); slept0 = sys_slept_ms(); mono.start();
        timer = new QTimer();
        const char *ms = getenv("VBRT_SYSINFO_POLL");
        timer->setInterval(ms && atoi(ms) > 0 ? atoi(ms) : 1000);
        QObject::connect(timer, &QTimer::timeout, [this]() { poll(); });
        timer->start();
#ifndef _WIN32
        sleep = new SysSleep;
        if (!sleep->open()) { delete sleep; sleep = nullptr; }
        else {
            sleeptimer = new QTimer();
            sleeptimer->setInterval(200);
            QObject::connect(sleeptimer, &QTimer::timeout, [this]() {
                int r;
                while (sleep && (r = sleep->next()) >= 0) {
                    if (r == 1) { going_to_sleep(); if (sleep) sleep->release(); }
                    else { slept0 = sys_slept_ms(); fire("powerresume"); if (sleep) sleep->inhibit(); }
                }
            });
            sleeptimer->start();
        }
#else
        native = make_power_filter(this);
        qApp->installNativeEventFilter(native);
#endif
    }
#ifdef _WIN32
    static QAbstractNativeEventFilter *make_power_filter(OcxSysInfo *si);
#endif
    void poll()
    {
        /* the clock: the wall time moved more than the time that passed */
        long long slept = sys_slept_ms();
        long long passed = mono.restart() + (slept - slept0);
        long long wall = QDateTime::currentMSecsSinceEpoch();
        bool resumed = slept - slept0 > 2000;
        if (resumed) {
#ifndef _WIN32
            if (!sleep) fire("powerresume");     /* without logind: from the clocks (Windows: WM_POWERBROADCAST) */
#endif
        } else if (llabs((wall - wall0) - passed) > 2000) fire("timechanged");
        wall0 = wall; slept0 = slept;
        SysPower p = sys_power();
        if (p.ac != power.ac || p.status != power.status) { power = p; fire("powerstatuschanged"); }
        else power = p;
        QStringList v = sys_volumes();
        if (v != vols) {
            QStringList old = vols;
            vols = v;
            for (const QString &x : v) if (!old.contains(x)) device("devicearrival", x);
            for (const QString &x : old) if (!v.contains(x)) device("deviceremovecomplete", x);
        }
    }
    QWidget *make(QWidget *) override { QTimer::singleShot(0, [this]() { start(); }); return nullptr; }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)val; (void)argc; (void)argv;
        *found = 1;
        QScreen *sc = QGuiApplication::primaryScreen();
        QRect wa = sc ? sc->availableGeometry() : QRect(0, 0, 1024, 768);
        if (set && n != "tag") { vb_raise(383); return vb_v_empty(); }
        if (n == "workareawidth") return vb_v_f((float)(wa.width() * 15));
        if (n == "workareaheight") return vb_v_f((float)(wa.height() * 15));
        if (n == "workarealeft") return vb_v_f((float)(wa.left() * 15));
        if (n == "workareatop") return vb_v_f((float)(wa.top() * 15));
        if (n == "scrollbarsize") return vb_v_f((float)(QApplication::style()->pixelMetric(QStyle::PM_ScrollBarExtent) * 15));
        if (n == "osplatform") return vb_v_i(2);           /* VER_PLATFORM_WIN32_NT */
        if (n == "osversion" || n == "osbuild") {
            int major = 6, minor = 1, build = 7601;     /* Linux: the Windows 7 that VB6 knows */
#ifdef _WIN32
            typedef LONG (WINAPI *rgv_t)(OSVERSIONINFOW *);
            if (rgv_t rgv = (rgv_t)(void *)GetProcAddress(GetModuleHandleW(L"ntdll.dll"), "RtlGetVersion")) {
                OSVERSIONINFOW vi; memset(&vi, 0, sizeof vi); vi.dwOSVersionInfoSize = sizeof vi;
                if (!rgv(&vi)) { major = (int)vi.dwMajorVersion; minor = (int)vi.dwMinorVersion; build = (int)vi.dwBuildNumber; }
            }
#endif
            if (n == "osbuild") return vb_v_i((vb_i16)build);
            return vb_v_f((float)(major + minor / 100.0));
        }
        SysPower p = sys_power();
        if (n == "acstatus") return vb_v_i((vb_i16)p.ac);
        if (n == "batterystatus") return vb_v_i((vb_i16)p.status);
        if (n == "batterylifepercent") return vb_v_i((vb_i16)p.percent);
        if (n == "batterylifetime") return vb_v_l(p.life);
        if (n == "batteryfulltime") return vb_v_l(p.full);
        *found = 0;
        return vb_v_empty();
    }
};
#ifdef _WIN32
/* the power messages of Windows: PBT_APMQUERYSUSPEND (0, Windows XP: Cancel refuses the sleep),
 * PBT_APMSUSPEND (4), PBT_APMRESUMESUSPEND (7) / PBT_APMRESUMEAUTOMATIC (0x12) / PBT_APMRESUMECRITICAL (6);
 * every top-level window gets them, the control tells each once */
struct SysPowerFilter : QAbstractNativeEventFilter {
    OcxSysInfo *si;
    unsigned long long last[32] = {};
    explicit SysPowerFilter(OcxSysInfo *s) : si(s) {}
#if QT_VERSION >= QT_VERSION_CHECK(6, 0, 0)
    bool nativeEventFilter(const QByteArray &type, void *message, qintptr *result) override
#else
    bool nativeEventFilter(const QByteArray &type, void *message, long *result) override
#endif
    {
        (void)type;
        MSG *m = (MSG *)message;
        if (!m || m->message != WM_POWERBROADCAST) return false;
        unsigned w = (unsigned)m->wParam;
        if (w >= 32) return false;
        unsigned long long now = GetTickCount64();
        bool again = last[w] && now - last[w] < 3000;
        if (w == 0) {                               /* the question: the answer of the first window counts */
            if (again) return false;
            last[w] = now;
            vbvar a[1] = { vb_v_bool(0) };
            si->fire("powerquerysuspend", 1, a);
            if (vb_v_truth(a[0])) { *result = 0x424D5144; return true; }   /* BROADCAST_QUERY_DENY */
            return false;
        }
        if (again) return false;
        last[w] = now;
        if (w == 4) si->fire("powersuspend");
        else if (w == 6 || w == 7 || w == 0x12) {
            if (!si->last_resume || now - si->last_resume > 5000) { si->last_resume = now; si->fire("powerresume"); }
        }
        return false;
    }
};
QAbstractNativeEventFilter *OcxSysInfo::make_power_filter(OcxSysInfo *si) { return new SysPowerFilter(si); }
#endif
OcxSysInfo::~OcxSysInfo()
{
    delete timer; delete sleeptimer;
#ifndef _WIN32
    delete sleep;
#else
    if (native) { qApp->removeNativeEventFilter(native); delete native; }
#endif
    if (watch) { qApp->removeEventFilter(watch); delete watch; }
}
bool SysWatch::eventFilter(QObject *o, QEvent *e)
{
    if (o == qApp && e->type() == QEvent::ApplicationPaletteChange) si->fire("syscolorschanged");
    return false;
}
static Ocx *new_sysinfo() { return new OcxSysInfo; }
