/* ocx_chart.inc - MSChart (MSCHRT20.OCX): the data grid, the object model and its own drawing.
 * The parts of the chart (Plot, Axis, Series, VtColor ...) keep their properties by path
 * ("plot.axis(1).valuescale.maximum"); ocx_chart_model.inc says what each member is.
 * Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "ocx_chart_model.inc"

static const ChartMember *chart_member(const char *iface, const std::string &n)
{
    for (const auto &m : chart_members) if (!strcmp(m.iface, iface) && n == m.name) return &m;
    return nullptr;
}
static const ChartIface *chart_iface(const char *iface)
{
    for (const auto &i : chart_ifaces) if (!strcmp(i.iface, iface)) return &i;
    return nullptr;
}
static void chart_error(vb_i32 n, const char *text)
{
    vb_i32 m = vb_mark();
    vb_raise_full(n, vs(QString("MSChart")), vs(QString(text)));
    vb_release(m);
}

struct OcxChart;
/* a part of the chart: Plot, Axis(1), SeriesCollection(2) ... */
struct ChartPart : OcxItem {
    vbobj ctl = nullptr;
    QString path;
    const char *iface = "";
    ~ChartPart() { if (ctl) vb_orelease(ctl); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};

class ChartWidget : public QWidget {
public:
    OcxChart *chart = nullptr;
    explicit ChartWidget(QWidget *parent) : QWidget(parent) { setAttribute(Qt::WA_OpaquePaintEvent); }
protected:
    void paintEvent(QPaintEvent *) override;
};

struct OcxChart : Ocx {
    QPointer<ChartWidget> w;
    int type = 1;                               /* VtChChartType2dBar */
    int rows = 0, cols = 0, row = 1, col = 1, rowlabix = 1, collabix = 1, seriescol = 1;
    std::vector<std::vector<double>> data;      /* [row][col] */
    std::vector<std::vector<bool>> nulls;
    std::vector<std::vector<QString>> rowlab, collab;    /* [level][row], [level][col] */
    bool legend = false, stacking = false, autoinc = false, repaint = true;
    bool allowsel = true, allowseries = true, allowrot = true, dither = true;
    QString title, footnote;
    QMap<QString, QVariant> bag;                /* the properties of the parts, by path */

    QWidget *make(QWidget *parent) override
    {
        w = new ChartWidget(parent);
        w->chart = this;
        defaults();
        return w;
    }
    void defaults()
    {
        type = 1; legend = false; stacking = false; autoinc = false;
        title.clear(); footnote.clear(); bag.clear();
        rowlab.assign(1, {}); collab.assign(1, {});
        set_size(1, 1, 5, 4);
        /* VB6: a new chart has random whole numbers 0 - 99 */
        for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) data[r][c] = (double)(std::rand() % 100);
        row = col = 1; rowlabix = collabix = 1;
    }
    void update() { if (w && repaint) w->update(); }
    /* the chart of the design: the settings of its OleObjectBlob as the compiler read them (lbvc/project/chartblob.py):
     * "#size" rowlabels,collabels,rows,cols, "#data(r,c)" (empty: no value), "#rowlabel(level,r)", "#collabel(level,c)",
     * the properties of the control (ChartType ...) and the parts by path ("plot.axis(1).valuescale.maximum") */
    void load(const char *const *p) override
    {
        for (int i = 0; p && p[i]; i += 2) {
            QString k = QString::fromUtf8(p[i]), v = QString::fromUtf8(p[i + 1] + 1);
            bool isstr = p[i + 1][0] == 's';
            double n = isstr ? 0 : v.toDouble();
            auto two = [&](int *a, int *b2) {
                int o = k.indexOf('('), c = k.indexOf(',');
                *a = k.mid(o + 1, c - o - 1).toInt(); *b2 = k.mid(c + 1, k.indexOf(')') - c - 1).toInt();
            };
            if (k == "#size") {
                QStringList q = v.split(',');
                if (q.size() == 4) {
                    rowlab.assign(1, {}); collab.assign(1, {});
                    set_size(q[0].toInt(), q[1].toInt(), q[2].toInt(), q[3].toInt());
                    for (auto &lv : rowlab) for (auto &t : lv) t.clear();
                    for (auto &lv : collab) for (auto &t : lv) t.clear();
                }
            } else if (k.startsWith("#data(")) {
                int r, c2; two(&r, &c2);
                if (r >= 1 && r <= rows && c2 >= 1 && c2 <= cols) {
                    nulls[r - 1][c2 - 1] = isstr;
                    data[r - 1][c2 - 1] = isstr ? 0 : n;
                }
            } else if (k.startsWith("#rowlabel(")) {
                int l, r; two(&l, &r);
                if (l >= 1 && l <= (int)rowlab.size() && r >= 1 && r <= rows) rowlab[l - 1][r - 1] = v;
            } else if (k.startsWith("#collabel(")) {
                int l, c2; two(&l, &c2);
                if (l >= 1 && l <= (int)collab.size() && c2 >= 1 && c2 <= cols) collab[l - 1][c2 - 1] = v;
            } else if (k == "ChartType") type = (int)n;
            else if (k == "ShowLegend") legend = n != 0;
            else if (k == "Stacking") stacking = n != 0;
            else if (k == "AutoIncrement") autoinc = n != 0;
            else if (k == "Repaint") repaint = n != 0;
            else if (k == "AllowSelections") allowsel = n != 0;
            else if (k == "AllowSeriesSelection") allowseries = n != 0;
            else if (k == "AllowDynamicRotation") allowrot = n != 0;
            else if (k == "AllowDithering") dither = n != 0;
            else if (k == "DoSetCursor") bag["dosetcursor"] = n != 0;
            else if (k == "TextLengthType" || k == "DrawMode") bag[k.toLower()] = (int)n;
            else if (k.contains('.')) bag[k] = isstr ? QVariant(v) : QVariant(n);
        }
        update();
    }
    /* SetSize: labels and data; new rows get labels R<n>, new columns C<n> */
    void set_size(int rlc, int clc, int nr, int nc)
    {
        rlc = std::max(1, rlc); clc = std::max(1, clc); nr = std::max(1, nr); nc = std::max(1, nc);
        data.resize(nr); nulls.resize(nr);
        for (int r = 0; r < nr; r++) { data[r].resize(nc, 0); nulls[r].resize(nc, false); }
        rows = nr; cols = nc;
        rowlab.resize(rlc); collab.resize(clc);
        for (int l = 0; l < rlc; l++) { int k = (int)rowlab[l].size(); rowlab[l].resize(nr); if (l == 0) for (int r = k; r < nr; r++) rowlab[0][r] = "R" + QString::number(r + 1); }
        for (int l = 0; l < clc; l++) { int k = (int)collab[l].size(); collab[l].resize(nc); if (l == 0) for (int c = k; c < nc; c++) collab[0][c] = "C" + QString::number(c + 1); }
        row = std::min(row, rows); col = std::min(col, cols);
        update();
    }
    static QString num(double v)
    {
        vb_i32 m = vb_mark();
        QString s = qs(vb_v2s(vb_v_d(v)));
        vb_release(m);
        return s;
    }
    vbvar chartdata()
    {
        vbarr *a = vb_atmp();
        /* (0 To RowCount, 0 To ColumnCount): the labels in row 0 and column 0 */
        vb_adim(a, sizeof(vbvar), VBA_VAR, 0, 2, 0, rows, 0, cols);
        if (vb_errpend) return vb_v_empty();
        for (int r = 0; r <= rows; r++)
            for (int c = 0; c <= cols; c++) {
                vbvar *e = (vbvar *)vb_aelem(a, 2, r, c);
                vbvar v = vb_v_empty();
                if (r == 0 && c > 0) v = vstr(collab[0][c - 1]);
                else if (c == 0 && r > 0) v = vstr(rowlab[0][r - 1]);
                else if (r > 0 && c > 0) v = nulls[r - 1][c - 1] ? vb_v_empty() : vb_v_d(data[r - 1][c - 1]);
                vb_vset(e, v);
            }
        return vb_v_arr(a, VT_VARIANT);
    }
    static vbvar elem(vbarr *a, vb_i32 evt, int i, int j)
    {
        void *p = a->nd == 1 ? vb_aelem(a, 1, i) : vb_aelem(a, 2, i, j);
        if (!p) return vb_v_empty();
        switch (evt) {
        case VT_VARIANT: return *(vbvar *)p;
        case VT_R8: return vb_v_d(*(double *)p);
        case VT_R4: return vb_v_f(*(float *)p);
        case VT_I2: return vb_v_i(*(vb_i16 *)p);
        case VT_I4: return vb_v_l(*(vb_i32 *)p);
        case VT_BOOL: return vb_v_bool(*(vb_i16 *)p);
        case VT_BSTR: return vb_v_s(*(vbs *)p);
        case VT_CY: return vb_v_c(*(vb_cy *)p);
        case VT_DATE: return vb_v_date(*(double *)p);
        case VT_UI1: return vb_v_byte(*(vb_u8 *)p);
        default: return vb_v_empty();
        }
    }
    /* ChartData = a: strings in the first row are the labels of the columns, in the first column those of the rows */
    void set_chartdata(vbvar v)
    {
        if (!(v.vt & VT_ARRAY)) { vb_raise(380); return; }
        vbarr *a = vb_v_arrptr(v);
        if (vb_errpend || !a || a->nd < 1 || a->nd > 2) { vb_raise(380); return; }
        vb_i32 evt = v.vt & ~VT_ARRAY;
        int r0 = a->lb[0], r1 = a->ub[0], c0 = a->nd == 2 ? a->lb[1] : 0, c1 = a->nd == 2 ? a->ub[1] : 0;
        if (r1 < r0 || c1 < c0) { vb_raise(380); return; }
        bool is_s = false;
        auto str = [&](int i, int j) { vbvar e = elem(a, evt, i, j); return e.vt == VT_BSTR; };
        bool collabels = false, rowlabels = false;
        for (int j = c0; j <= c1; j++) if (str(r0, j)) collabels = true;
        for (int i = r0; i <= r1; i++) if (str(i, c0)) rowlabels = true;
        if (a->nd == 1) collabels = false;
        (void)is_s;
        int dr0 = r0 + (collabels ? 1 : 0), dc0 = c0 + (rowlabels ? 1 : 0);
        int nr = std::max(1, r1 - dr0 + 1), nc = std::max(1, c1 - dc0 + 1);
        rowlab.assign(1, {}); collab.assign(1, {});
        data.clear(); nulls.clear();
        set_size(1, 1, nr, nc);
        for (int i = dr0; i <= r1; i++)
            for (int j = dc0; j <= c1; j++) {
                vbvar e = elem(a, evt, i, j);
                int r = i - dr0, c = j - dc0;
                if (e.vt == VT_EMPTY || e.vt == VT_NULL) { nulls[r][c] = true; data[r][c] = 0; continue; }
                vb_i32 m = vb_mark();
                double d = vb_v2d(e);
                if (vb_errpend) { vb_err_clear(); d = 0; }
                vb_release(m);
                data[r][c] = d; nulls[r][c] = false;
            }
        if (collabels) for (int j = dc0; j <= c1; j++) { vbvar e = elem(a, evt, r0, j); collab[0][j - dc0] = e.vt == VT_BSTR ? qs(e.u.s) : QString(); }
        if (rowlabels) for (int i = dr0; i <= r1; i++) { vbvar e = elem(a, evt, i, c0); rowlab[0][i - dr0] = e.vt == VT_BSTR ? qs(e.u.s) : QString(); }
        row = col = 1;
        update();
    }
    bool check_row(int r) { if (r < 1 || r > rows) { chart_error(1117, "Invalid index"); return false; } return true; }
    bool check_col(int c) { if (c < 1 || c > cols) { chart_error(1117, "Invalid index"); return false; } return true; }
    vbobj part(const QString &path, const char *iface)
    {
        auto *p = new ChartPart;
        p->ctl = c->self; vb_oaddref(p->ctl);
        p->path = path; p->iface = iface;
        const ChartIface *ci = chart_iface(iface);
        p->obj = vb_onew(ocx_class(iface, ci && *ci->defmember ? ci->defmember : nullptr, false));     /* TypeName: IVcPlot */
        p->obj->ext = p;
        return p->obj;
    }
    /* the color of a part, if the program set one (VtColor.Set r, g, b or .Red ...) */
    bool color(const QString &path, QColor *out)
    {
        if (!bag.contains(path + ".red") && !bag.contains(path + ".green") && !bag.contains(path + ".blue")) return false;
        if (bag.value(path + ".automatic").toBool()) return false;
        *out = QColor(bag.value(path + ".red").toInt() & 255, bag.value(path + ".green").toInt() & 255, bag.value(path + ".blue").toInt() & 255);
        return true;
    }
    QColor series_color(int c)
    {
        static const QColor pal[] = { QColor(0, 0, 255), QColor(255, 0, 0), QColor(0, 255, 0), QColor(255, 255, 0),
                                      QColor(255, 0, 255), QColor(0, 255, 255), QColor(0, 0, 128), QColor(128, 0, 0),
                                      QColor(0, 128, 0), QColor(128, 128, 0), QColor(128, 0, 128), QColor(0, 128, 128) };
        QString base = QString("plot.seriescollection(%1)").arg(c + 1);
        QColor q;
        if (color(base + ".datapoints(-1).brush.fillcolor", &q)) return q;
        if (color(base + ".pen.vtcolor", &q)) return q;
        return pal[c % 12];
    }
    bool hidden(int c) { QString b = QString("plot.seriescollection(%1).position.").arg(c + 1); return bag.value(b + "hidden").toBool() || bag.value(b + "excluded").toBool(); }
    QString rowtext(int r)
    {
        QStringList parts;
        for (int l = (int)rowlab.size() - 1; l >= 0; l--) if (r < (int)rowlab[l].size() && !rowlab[l][r].isEmpty()) parts << rowlab[l][r];
        return parts.join(' ');
    }
    QString coltext(int c)
    {
        QString lt = bag.value(QString("plot.seriescollection(%1).legendtext").arg(c + 1)).toString();
        if (!lt.isEmpty()) return lt;
        QStringList parts;
        for (int l = (int)collab.size() - 1; l >= 0; l--) if (c < (int)collab[l].size() && !collab[l][c].isEmpty()) parts << collab[l][c];
        return parts.join(' ');
    }
    void paint(QPainter &p, const QRect &all);
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
    vbvar part_member(ChartPart *pt, const std::string &n, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found);
};
void ChartWidget::paintEvent(QPaintEvent *)
{
    QPainter p(this);
    if (chart) chart->paint(p, rect());
}

/* ---------------------------------------------------------------- the drawing */
static double nice_step(double range)
{
    if (range <= 0) return 1;
    double raw = range / 5, mag = std::pow(10, std::floor(std::log10(raw))), f = raw / mag;
    return (f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10) * mag;
}
void OcxChart::paint(QPainter &p, const QRect &all)
{
    QColor back(255, 255, 255);
    color("backdrop.fill.brush.fillcolor", &back);
    p.fillRect(all, back);
    p.setRenderHint(QPainter::Antialiasing, true);
    QFont f = c && c->w ? c->w->font() : QFont();
    p.setFont(f);
    QFontMetrics fm(f);
    QRect r = all.adjusted(6, 6, -6, -6);
    QString tt = bag.value("title.text").toString();
    if (!tt.isEmpty()) {
        QFont tf = f; tf.setPointSizeF(f.pointSizeF() * 1.4); tf.setBold(true);
        p.setFont(tf);
        QFontMetrics tm(tf);
        p.setPen(Qt::black);
        p.drawText(QRect(r.left(), r.top(), r.width(), tm.height()), Qt::AlignHCenter, tt);
        r.setTop(r.top() + tm.height() + 4);
        p.setFont(f);
    }
    QString ft = bag.value("footnote.text").toString();
    if (!ft.isEmpty()) {
        p.setPen(Qt::black);
        p.drawText(QRect(r.left(), r.bottom() - fm.height(), r.width(), fm.height()), Qt::AlignHCenter, ft);
        r.setBottom(r.bottom() - fm.height() - 4);
    }
    std::vector<int> vis;
    for (int k = 0; k < cols; k++) if (!hidden(k)) vis.push_back(k);
    bool pie = type == 14 || type == 13, xy = type == 16;
    if (legend) {
        int lw = 0;
        int nitems = pie ? cols : (int)vis.size();
        for (int k = 0; k < nitems; k++) lw = std::max(lw, fm.horizontalAdvance(coltext(pie ? k : vis[k])));
        lw += fm.height() + 12;
        int lh = nitems * (fm.height() + 2);
        QRect lr(r.right() - lw, r.center().y() - lh / 2, lw, lh);
        for (int k = 0; k < nitems; k++) {
            int s = pie ? k : vis[k];
            QRect box(lr.left() + 2, lr.top() + k * (fm.height() + 2) + 2, fm.height() - 4, fm.height() - 4);
            p.fillRect(box, series_color(s)); p.setPen(Qt::black); p.drawRect(box);
            p.drawText(QPoint(box.right() + 6, box.top() + fm.ascent() - 2), coltext(s));
        }
        r.setRight(lr.left() - 8);
    }
    if (rows < 1 || cols < 1 || r.width() < 20 || r.height() < 20) return;
    p.setPen(Qt::black);
    if (pie) {                                  /* a pie for each row, a slice for each column */
        int n = rows, side = std::min(r.width() / n, r.height() - fm.height() - 4);
        for (int i = 0; i < n; i++) {
            double tot = 0;
            for (int k = 0; k < cols; k++) if (!nulls[i][k]) tot += std::fabs(data[i][k]);
            QRect pr(r.left() + i * (r.width() / n) + (r.width() / n - side) / 2, r.top(), side, side);
            int a = 90 * 16;
            for (int k = 0; k < cols && tot > 0; k++) {
                int span = -(int)std::lround(std::fabs(data[i][k]) / tot * 360 * 16);
                p.setBrush(series_color(k));
                p.drawPie(pr, a, span);
                a += span;
            }
            p.drawText(QRect(pr.left(), pr.bottom() + 2, pr.width(), fm.height()), Qt::AlignHCenter, rowtext(i));
        }
        p.setBrush(Qt::NoBrush);
        return;
    }
    /* the value axis */
    double lo = 0, hi = 0;
    bool any = false;
    auto see = [&](double v) { if (!any) { lo = hi = v; any = true; } lo = std::min(lo, v); hi = std::max(hi, v); };
    bool stack = stacking && (type == 0 || type == 1 || type == 4 || type == 5);
    for (int i = 0; i < rows; i++) {
        double pos = 0, neg = 0;
        for (size_t q = 0; q < vis.size(); q++) {
            int k = vis[q];
            if (xy && q % 2 == 0) continue;
            if (nulls[i][k]) continue;
            if (stack) { if (data[i][k] >= 0) pos += data[i][k]; else neg += data[i][k]; see(pos); see(neg); }
            else see(data[i][k]);
        }
    }
    lo = std::min(lo, 0.0); hi = std::max(hi, 0.0);
    if (hi == lo) hi = lo + 1;
    double step = nice_step(hi - lo);
    double vmin = std::floor(lo / step) * step, vmax = std::ceil(hi / step) * step;
    if (bag.contains("plot.axis(1).valuescale.auto") && !bag.value("plot.axis(1).valuescale.auto").toBool()) {
        if (bag.contains("plot.axis(1).valuescale.minimum")) vmin = bag.value("plot.axis(1).valuescale.minimum").toDouble();
        if (bag.contains("plot.axis(1).valuescale.maximum")) vmax = bag.value("plot.axis(1).valuescale.maximum").toDouble();
        int md = bag.value("plot.axis(1).valuescale.majordivision").toInt();
        if (md > 0) step = (vmax - vmin) / md;
    }
    if (vmax <= vmin) vmax = vmin + 1;
    int lab = 0;
    for (double v = vmin; v <= vmax + step / 2; v += step) lab = std::max(lab, fm.horizontalAdvance(num(v)));
    QRect pa(r.left() + lab + 8, r.top() + fm.height() / 2, r.width() - lab - 12, r.height() - fm.height() * 2 - 4);
    if (pa.width() < 10 || pa.height() < 10) return;
    auto ypos = [&](double v) { return pa.bottom() - (v - vmin) / (vmax - vmin) * pa.height(); };
    p.setPen(QColor(192, 192, 192));
    for (double v = vmin; v <= vmax + step / 2; v += step) {
        double y = ypos(v);
        p.drawLine(QPointF(pa.left(), y), QPointF(pa.right(), y));
    }
    p.setPen(Qt::black);
    for (double v = vmin; v <= vmax + step / 2; v += step) {
        double y = ypos(v);
        p.drawText(QRectF(r.left(), y - fm.height() / 2.0, lab + 4, fm.height()), Qt::AlignRight | Qt::AlignVCenter, num(v));
    }
    p.drawLine(pa.bottomLeft(), pa.bottomRight());
    p.drawLine(pa.topLeft(), pa.bottomLeft());
    double zero = ypos(std::max(vmin, std::min(vmax, 0.0)));
    p.save();
    p.setClipRect(pa.adjusted(0, -1, 1, 1));
    if (xy) {                                   /* XY: the columns in pairs (x, y) */
        double xlo = 0, xhi = 1; bool xa = false;
        for (int i = 0; i < rows; i++) for (size_t q = 0; q + 1 < vis.size(); q += 2) {
            double x = data[i][vis[q]]; if (!xa) { xlo = xhi = x; xa = true; } xlo = std::min(xlo, x); xhi = std::max(xhi, x);
        }
        if (xhi == xlo) xhi = xlo + 1;
        for (size_t q = 0; q + 1 < vis.size(); q += 2) {
            QPolygonF pl;
            for (int i = 0; i < rows; i++)
                pl << QPointF(pa.left() + (data[i][vis[q]] - xlo) / (xhi - xlo) * pa.width(), ypos(data[i][vis[q + 1]]));
            p.setPen(QPen(series_color(vis[q + 1]), 2)); p.drawPolyline(pl);
            for (auto &pt : pl) p.drawEllipse(pt, 2.5, 2.5);
        }
        p.restore();
        return;
    }
    double gw = (double)pa.width() / rows;
    int nv = std::max(1, (int)vis.size());
    if (type == 0 || type == 1 || type == 8 || type == 9 || type == 10 || type == 11 || type == 12) {    /* bars */
        for (int i = 0; i < rows; i++) {
            double pos = 0, neg = 0;
            double bw = stack ? gw * 0.6 : gw * 0.8 / nv;
            for (int q = 0; q < (int)vis.size(); q++) {
                int k = vis[q];
                if (nulls[i][k]) continue;
                double v = data[i][k], b0, b1;
                if (stack) { if (v >= 0) { b0 = pos; pos += v; b1 = pos; } else { b0 = neg; neg += v; b1 = neg; } }
                else { b0 = 0; b1 = v; }
                double x = pa.left() + i * gw + (stack ? gw * 0.2 : gw * 0.1 + q * bw);
                QRectF br(QPointF(x, ypos(std::max(b0, b1))), QPointF(x + bw, ypos(std::min(b0, b1))));
                p.setPen(Qt::black); p.setBrush(series_color(k)); p.drawRect(br);
            }
        }
    } else {                                    /* lines, areas, steps */
        std::vector<double> base(rows, 0.0);
        for (int q = 0; q < (int)vis.size(); q++) {
            int k = vis[q];
            QPolygonF pl;
            std::vector<double> top(rows);
            for (int i = 0; i < rows; i++) {
                top[i] = (stack ? base[i] : 0) + (nulls[i][k] ? 0 : data[i][k]);
                double x = pa.left() + (i + 0.5) * gw;
                if (type == 6 || type == 7) { pl << QPointF(pa.left() + i * gw, ypos(top[i])) << QPointF(pa.left() + (i + 1) * gw, ypos(top[i])); }
                else pl << QPointF(x, ypos(top[i]));
            }
            QColor sc = series_color(k);
            if (type == 4 || type == 5) {
                QPolygonF area = pl;
                for (int i = rows - 1; i >= 0; i--) area << QPointF(pa.left() + (i + 0.5) * gw, stack ? ypos(base[i]) : zero);
                p.setPen(Qt::black); p.setBrush(sc); p.drawPolygon(area);
            } else {
                p.setBrush(Qt::NoBrush); p.setPen(QPen(sc, 2)); p.drawPolyline(pl);
            }
            if (stack) base = top;
        }
    }
    p.restore();
    p.setBrush(Qt::NoBrush);
    p.setPen(Qt::black);
    for (int i = 0; i < rows; i++)
        p.drawText(QRectF(pa.left() + i * gw, pa.bottom() + 2, gw, fm.height()), Qt::AlignHCenter | Qt::AlignTop, rowtext(i));
}

/* ---------------------------------------------------------------- the members */
vbvar OcxChart::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    auto ival = [&]() { vb_i32 v = vb_v2l(val); return v; };
    if (n == "charttype") {
        if (!set) return vb_v_l(type);
        int v = ival(); if (vb_errpend) return vb_v_empty();
        if (v < 0 || v > 26) { vb_raise(380); return vb_v_empty(); }
        type = v; update(); return vb_v_empty();
    }
    if (n == "chartdata") { if (!set) return chartdata(); set_chartdata(val); return vb_v_empty(); }
    if (n == "rowcount") { if (!set) return vb_v_i((vb_i16)rows); int v = ival(); if (vb_errpend) return vb_v_empty(); if (v < 1) { vb_raise(380); return vb_v_empty(); } set_size((int)rowlab.size(), (int)collab.size(), v, cols); return vb_v_empty(); }
    if (n == "columncount") { if (!set) return vb_v_i((vb_i16)cols); int v = ival(); if (vb_errpend) return vb_v_empty(); if (v < 1) { vb_raise(380); return vb_v_empty(); } set_size((int)rowlab.size(), (int)collab.size(), rows, v); return vb_v_empty(); }
    if (n == "row") { if (!set) return vb_v_i((vb_i16)row); int v = ival(); if (vb_errpend || !check_row(v)) return vb_v_empty(); row = v; return vb_v_empty(); }
    if (n == "column") { if (!set) return vb_v_i((vb_i16)col); int v = ival(); if (vb_errpend || !check_col(v)) return vb_v_empty(); col = v; return vb_v_empty(); }
    if (n == "data") {
        if (!set) return vstr(nulls[row - 1][col - 1] ? QString() : num(data[row - 1][col - 1]));
        QString s = qs(vb_v2s(val)).trimmed();
        if (s.isEmpty()) nulls[row - 1][col - 1] = true;
        else {
            double d = vb_v2d(val); if (vb_errpend) return vb_v_empty();
            data[row - 1][col - 1] = d; nulls[row - 1][col - 1] = false;
        }
        if (autoinc) { if (row < rows) row++; else { row = 1; if (col < cols) col++; } }     /* down the column */
        update();
        return vb_v_empty();
    }
    if (n == "rowlabel") { int l = rowlabix - 1; if (!set) return vstr(l < (int)rowlab.size() ? rowlab[l][row - 1] : QString()); if (l < (int)rowlab.size()) rowlab[l][row - 1] = qs(vb_v2s(val)); update(); return vb_v_empty(); }
    if (n == "columnlabel") { int l = collabix - 1; if (!set) return vstr(l < (int)collab.size() ? collab[l][col - 1] : QString()); if (l < (int)collab.size()) collab[l][col - 1] = qs(vb_v2s(val)); update(); return vb_v_empty(); }
    if (n == "rowlabelcount") { if (!set) return vb_v_i((vb_i16)rowlab.size()); int v = ival(); if (vb_errpend) return vb_v_empty(); if (v < 0) { vb_raise(380); return vb_v_empty(); } set_size(v, (int)collab.size(), rows, cols); return vb_v_empty(); }
    if (n == "columnlabelcount") { if (!set) return vb_v_i((vb_i16)collab.size()); int v = ival(); if (vb_errpend) return vb_v_empty(); if (v < 0) { vb_raise(380); return vb_v_empty(); } set_size((int)rowlab.size(), v, rows, cols); return vb_v_empty(); }
    if (n == "rowlabelindex") { if (!set) return vb_v_i((vb_i16)rowlabix); int v = ival(); if (vb_errpend) return vb_v_empty(); if (v < 1 || v > (int)rowlab.size()) { chart_error(1117, "Invalid index"); return vb_v_empty(); } rowlabix = v; return vb_v_empty(); }
    if (n == "columnlabelindex") { if (!set) return vb_v_i((vb_i16)collabix); int v = ival(); if (vb_errpend) return vb_v_empty(); if (v < 1 || v > (int)collab.size()) { chart_error(1117, "Invalid index"); return vb_v_empty(); } collabix = v; return vb_v_empty(); }
    if (n == "titletext") { if (!set) return vstr(bag.value("title.text").toString()); bag["title.text"] = qs(vb_v2s(val)); update(); return vb_v_empty(); }
    if (n == "footnotetext") { if (!set) return vstr(bag.value("footnote.text").toString()); bag["footnote.text"] = qs(vb_v2s(val)); update(); return vb_v_empty(); }
    if (n == "showlegend") { if (!set) return vbool(legend); legend = vb_v2bool(val) != 0; update(); return vb_v_empty(); }
    if (n == "stacking") { if (!set) return vbool(stacking); stacking = vb_v2bool(val) != 0; update(); return vb_v_empty(); }
    if (n == "autoincrement") { if (!set) return vbool(autoinc); autoinc = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "repaint") { if (!set) return vbool(repaint); repaint = vb_v2bool(val) != 0; update(); return vb_v_empty(); }
    if (n == "randomfill") {
        if (!set) return vbool(false);
        if (vb_v2bool(val)) for (int r = 0; r < rows; r++) for (int k = 0; k < cols; k++) { data[r][k] = (double)(rand() % 100); nulls[r][k] = false; }
        update(); return vb_v_empty();
    }
    if (n == "allowselections") { if (!set) return vbool(allowsel); allowsel = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "allowseriesselection") { if (!set) return vbool(allowseries); allowseries = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "allowdynamicrotation") { if (!set) return vbool(allowrot); allowrot = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "allowdithering") { if (!set) return vbool(dither); dither = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "chart3d") { if (set) { vb_raise(383); return vb_v_empty(); } return vbool(type % 2 == 0 && type <= 8); }
    if (n == "activeseriescount") { if (set) { vb_raise(383); return vb_v_empty(); } int k = 0; for (int i = 0; i < cols; i++) if (!hidden(i)) k++; return vb_v_i((vb_i16)k); }
    if (n == "seriescolumn") { if (!set) return vb_v_i((vb_i16)seriescol); seriescol = ival(); return vb_v_empty(); }
    if (n == "seriestype") { QString k = QString("plot.seriescollection(%1).seriestype").arg(col); if (!set) return vb_v_l(bag.value(k).toInt()); bag[k] = (int)ival(); update(); return vb_v_empty(); }
    if (n == "textlengthtype" || n == "drawmode" || n == "dosetcursor" || n == "borderstyle") {
        QString k = QString::fromStdString(n);
        if (!set) return n == "dosetcursor" ? vbool(bag.value(k, true).toBool()) : vb_v_l(bag.value(k).toInt());
        bag[k] = n == "dosetcursor" ? QVariant(vb_v2bool(val) != 0) : QVariant((int)ival()); return vb_v_empty();
    }
    if (n == "todefaults") { defaults(); update(); return vb_v_empty(); }
    if (n == "refresh" || n == "layout") { if (w) w->repaint(); return vb_v_empty(); }
    if (n == "editcopy" || n == "editpaste" || n == "selectpart" || n == "aboutbox") return vb_v_empty();
    if (n == "chart") return vb_v_l(0);
    /* the parts */
    static const struct { const char *name, *iface; } parts[] = {
        { "title", "IVcTitle" }, { "footnote", "IVcFootnote" }, { "legend", "IVcLegend" }, { "plot", "IVcPlot" },
        { "datagrid", "IVcDataGrid" }, { "backdrop", "IVcBackdrop" } };
    for (auto &pt : parts) if (n == pt.name) { if (set) { vb_raise(383); return vb_v_empty(); } return retobj(part(QString(pt.name), pt.iface)); }
    (void)argc; (void)argv;
    *found = 0;
    return vb_v_empty();
}
/* Axis(1), SeriesCollection(2): the arguments become part of the path */
static QString path_args(vb_i32 argc, vbvar *argv)
{
    QStringList a;
    for (int i = 0; i < argc; i++) {
        if (vb_ismissing(argv[i])) continue;
        vb_i32 m = vb_mark();
        a << qs(vb_v2s(argv[i])).toLower();
        vb_release(m);
    }
    return a.isEmpty() ? QString() : "(" + a.join(',') + ")";
}
vbvar OcxChart::part_member(ChartPart *pt, const std::string &n, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    std::string ifc = pt->iface;
    QString base = pt->path;
    vb_i32 nargs = set ? argc - 1 : argc;
    auto iarg = [&](int i) { return i < nargs ? (int)vb_v2l(argv[i]) : 0; };
    /* the parts that are the data grid or follow the control's own properties */
    if (ifc == "IVcDataGrid") {
        if (n == "rowcount") { if (!set) return vb_v_i((vb_i16)rows); set_size((int)rowlab.size(), (int)collab.size(), (int)vb_v2l(val), cols); return vb_v_empty(); }
        if (n == "columncount") { if (!set) return vb_v_i((vb_i16)cols); set_size((int)rowlab.size(), (int)collab.size(), rows, (int)vb_v2l(val)); return vb_v_empty(); }
        if (n == "rowlabelcount") { if (!set) return vb_v_i((vb_i16)rowlab.size()); set_size((int)vb_v2l(val), (int)collab.size(), rows, cols); return vb_v_empty(); }
        if (n == "columnlabelcount") { if (!set) return vb_v_i((vb_i16)collab.size()); set_size((int)rowlab.size(), (int)vb_v2l(val), rows, cols); return vb_v_empty(); }
        if (n == "setsize") { set_size(iarg(0), iarg(1), iarg(2), iarg(3)); return vb_v_empty(); }
        if (n == "setdata") {
            int r = iarg(0), k = iarg(1);
            if (vb_errpend || !check_row(r) || !check_col(k)) return vb_v_empty();
            data[r - 1][k - 1] = nargs > 2 ? vb_v2d(argv[2]) : 0; nulls[r - 1][k - 1] = nargs > 3 && vb_v2l(argv[3]) != 0;
            update(); return vb_v_empty();
        }
        if (n == "getdata") return vb_v_empty();         /* its results are ByRef arguments */
        if (n == "rowlabel" || n == "columnlabel") {
            bool isrow = n == "rowlabel";
            int i = iarg(0), l = nargs > 1 ? iarg(1) : 1;
            auto &lab = isrow ? rowlab : collab;
            int cnt = isrow ? rows : cols;
            if (vb_errpend) return vb_v_empty();
            if (i < 1 || i > cnt || l < 1 || l > (int)lab.size()) { chart_error(1117, "Invalid index"); return vb_v_empty(); }
            if (!set) return vstr(lab[l - 1][i - 1]);
            lab[l - 1][i - 1] = qs(vb_v2s(val)); update(); return vb_v_empty();
        }
        if (n == "compositerowlabel") { int i = iarg(0); if (i < 1 || i > rows) { chart_error(1117, "Invalid index"); return vb_v_empty(); } return vstr(rowtext(i - 1)); }
        if (n == "compositecolumnlabel") { int i = iarg(0); if (i < 1 || i > cols) { chart_error(1117, "Invalid index"); return vb_v_empty(); } QStringList ps; for (int l = (int)collab.size() - 1; l >= 0; l--) if (!collab[l][i - 1].isEmpty()) ps << collab[l][i - 1]; return vstr(ps.join(' ')); }
        if (n == "insertrows" || n == "deleterows" || n == "insertcolumns" || n == "deletecolumns") {
            int at = iarg(0) - 1, cnt = iarg(1);
            bool isrow = n.find("rows") != std::string::npos, ins = n[0] == 'i';
            int total = isrow ? rows : cols;
            if (vb_errpend) return vb_v_empty();
            if (at < 0 || at > total - (ins ? 0 : 1) || cnt < 1 || (!ins && at + cnt > total)) { chart_error(1117, "Invalid index"); return vb_v_empty(); }
            for (int t = 0; t < cnt; t++) {
                if (isrow) {
                    if (ins) { data.insert(data.begin() + at, std::vector<double>(cols, 0)); nulls.insert(nulls.begin() + at, std::vector<bool>(cols, false)); for (size_t l = 0; l < rowlab.size(); l++) rowlab[l].insert(rowlab[l].begin() + at, l ? QString() : "R" + QString::number(at + 1)); }
                    else { data.erase(data.begin() + at); nulls.erase(nulls.begin() + at); for (auto &lv : rowlab) lv.erase(lv.begin() + at); }
                } else {
                    for (int r = 0; r < rows; r++) {
                        if (ins) { data[r].insert(data[r].begin() + at, 0); nulls[r].insert(nulls[r].begin() + at, false); }
                        else { data[r].erase(data[r].begin() + at); nulls[r].erase(nulls[r].begin() + at); }
                    }
                    for (size_t l = 0; l < collab.size(); l++) { if (ins) collab[l].insert(collab[l].begin() + at, l ? QString() : "C" + QString::number(at + 1)); else collab[l].erase(collab[l].begin() + at); }
                }
            }
            if (isrow) rows = (int)data.size(); else cols = (int)collab[0].size();
            row = std::max(1, std::min(row, rows)); col = std::max(1, std::min(col, cols));
            update(); return vb_v_empty();
        }
        if (n == "randomdatafill" || n == "randomfillcolumns" || n == "randomfillrows") {
            for (int r = 0; r < rows; r++) for (int k = 0; k < cols; k++) { data[r][k] = (double)(rand() % 100); nulls[r][k] = false; }
            update(); return vb_v_empty();
        }
        if (n == "initializelabels") { for (int r = 0; r < rows; r++) rowlab[0][r] = "R" + QString::number(r + 1); for (int k = 0; k < cols; k++) collab[0][k] = "C" + QString::number(k + 1); update(); return vb_v_empty(); }
    }
    if (ifc == "IVcSeriesCollection" && n == "count") return vb_v_l(cols);
    if (ifc == "IVcDataPoints" && n == "count") return vb_v_l(rows);
    if (ifc == "IVcLegend" && n == "location") { /* Legend.Location.Visible follows ShowLegend */ }
    if (ifc == "IVcLocation" && n == "visible" && base == "legend.location") { if (!set) return vbool(legend); legend = vb_v2bool(val) != 0; update(); return vb_v_empty(); }
    if (ifc == "IVcColor" && n == "set") {
        bag[base + ".red"] = iarg(0); bag[base + ".green"] = iarg(1); bag[base + ".blue"] = iarg(2);
        bag[base + ".automatic"] = false;
        update(); return vb_v_empty();
    }
    const ChartMember *m = chart_member(pt->iface, n);
    if (!m) { *found = 0; return vb_v_empty(); }
    QString key = base + "." + QString::fromStdString(n) + path_args(nargs, argv);
    if (m->kind == 'o') {
        if (set) { vb_raise(383); return vb_v_empty(); }
        if (ifc == "IVcSeriesCollection" || ifc == "IVcDataPoints") {
            int i = iarg(0);
            int cnt = ifc == "IVcSeriesCollection" ? cols : rows;
            if (vb_errpend) return vb_v_empty();
            if (!(ifc == "IVcDataPoints" && i == -1) && (i < 1 || i > cnt)) { chart_error(1117, "Invalid index"); return vb_v_empty(); }
            key = base + "(" + QString::number(i) + ")";
        } else if (m->args == 0 && nargs > 0) {
            /* Plot.SeriesCollection(1): the arguments go to the default member of the part (Item) */
            key = base + "." + QString::fromStdString(n);
            vbobj o = part(key, m->ret);
            vb_oaddref(o);
            vbvar r = vb_late_o(o, "", VBM_GET, nargs, argv);
            vb_orelease(o);
            return r;
        }
        return retobj(part(key, m->ret));
    }
    if (m->kind == 'm') return vb_v_empty();
    if (set) {
        if (!m->let) { vb_raise(383); return vb_v_empty(); }
        QVariant v;
        switch (m->kind) {
        case 'b': v = vb_v2bool(val) != 0; break;
        case 'i': v = (int)vb_v2l(val); break;
        case 'f': v = vb_v2d(val); break;
        default: v = qs(vb_v2s(val)); break;
        }
        if (vb_errpend) return vb_v_empty();
        bag[key] = v;
        if (ifc == "IVcColor" && n != "automatic") bag[base + ".automatic"] = false;
        update();
        return vb_v_empty();
    }
    QVariant v = bag.value(key);
    switch (m->kind) {
    case 'b': return vbool(v.isValid() ? v.toBool() : n == "auto" || n == "automatic" || n == "visible" || n == "show");
    case 'i': return vb_v_l(v.toInt());
    case 'f': return vb_v_d(v.toDouble());
    case 's': return vstr(v.toString());
    default: return v.isValid() ? vstr(v.toString()) : vb_v_empty();
    }
}
vbvar ChartPart::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    Ctl *k = ctl ? C(ctl) : nullptr;
    auto *ch = k ? dynamic_cast<OcxChart *>(k->ocx) : nullptr;
    if (!ch) { ocx_error(35605); return vb_v_empty(); }
    return ch->part_member(this, n, set, val, argc, argv, found);
}
