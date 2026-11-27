/* vbforms.cpp - forms and the intrinsic controls of VB6 on Qt6 (libvbforms).
 *
 * A form object is an object of the class the compiler makes from the .frm; its
 * built-in members (Caption, Show, the controls by name ...) come from here through
 * vb_class.fallback. Controls are objects of the classes below. Events go to the
 * handlers of the form ("command1_click", "form_load" ...) through vbf_formdef.events.
 *
 * Sizes are in twips like VB6: 15 twips = 1 pixel (96 DPI). Colors are VB6 colors:
 * &HBBGGRR, or &H800000nn for a system color.
 *
 * What the order of events follows (VB6 oracle, tests/forms):
 *   f_basic  Load/Activate/QueryUnload/Unload/Terminate, control arrays, Forms
 *   f_ui     graphics methods and scales, colors, fonts, focus, keyboard, SendKeys, menus
 *
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
extern "C" {
#include "../vbrt.h"
}
#include <QApplication>
#include <QWidget>
#include <QFrame>
#include <QPushButton>
#include <QLabel>
#include <QLineEdit>
#include <QPlainTextEdit>
#include <QTextCursor>
#include <QTextEdit>
#include <QTextList>
#include <QTextBlock>
#include <QTextDocumentFragment>
#include <QTextLayout>
#include <QAbstractTextDocumentLayout>
#include <QFile>
#include <QSet>
#include <QtEndian>
#include <QCheckBox>
#include <QRadioButton>
#include <QListWidget>
#include <QComboBox>
#include <QGroupBox>
#include <QScrollBar>
#include <QMessageBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QDir>
#include <QFileInfo>
#include <QStorageInfo>
#include <QMouseEvent>
#include <QGridLayout>
#include <QMdiArea>
#include <QMdiSubWindow>
#include <QClipboard>
#include <QMimeData>
#include <QScreen>
#include <QFontDatabase>
#include <QDateTime>
#include <QPdfWriter>
#include <QPageSize>
#include <QStyle>
#include <QStyleFactory>
#include <QPalette>
#include <QSettings>
#include <QProcess>
#include <QStyleHints>
#include <QGuiApplication>
#include <QMenuBar>
#include <QMenu>
#include <QAction>
#include <QVBoxLayout>
#include <QTimer>
#include <QEventLoop>
#include <QCloseEvent>
#include <QResizeEvent>
#include <QKeyEvent>
#include <QPainter>
#include <QImage>
#include <QPointer>
#include <QFontMetrics>
#include <QProgressBar>
#include <QSlider>
#include <QToolButton>
#include <QTabBar>
#include <QHBoxLayout>
#include <QToolTip>
#include <QStyleOptionSlider>
#include <QHelpEvent>
#include <QTreeWidget>
#include <QHeaderView>
#include <QListWidget>
#include <QStackedWidget>
#include <QTableWidget>
#include <QDateTimeEdit>
#include <QCalendarWidget>
#include <QFileDialog>
#include <QColorDialog>
#include <QFontDialog>

/* Qt 6 on Linux; Qt 5.15 for Windows programs (Qt 6 needs Windows 10, Qt 5.15 runs on Windows 7) */
#if QT_VERSION < QT_VERSION_CHECK(6, 0, 0)
static inline QPointF ev_pos(const QMouseEvent *e) { return e->localPos(); }
static inline QPointF ev_scenepos(const QMouseEvent *e) { return e->windowPos(); }
static inline QPointF ev_gpos(const QMouseEvent *e) { return e->screenPos(); }
static inline QPointF map_from_global(const QWidget *w, QPointF p) { return QPointF(w->mapFromGlobal(p.toPoint())); }
static inline QStringList font_families() { static QStringList f = QFontDatabase().families(); return f; }   /* (once) */
#else
static inline QPointF ev_pos(const QMouseEvent *e) { return e->position(); }
static inline QPointF ev_scenepos(const QMouseEvent *e) { return e->scenePosition(); }
static inline QPointF ev_gpos(const QMouseEvent *e) { return e->globalPosition(); }
static inline QPointF map_from_global(const QWidget *w, QPointF p) { return w->mapFromGlobal(p); }
/* the fonts of the system, read once: Screen.Fonts(i) in a loop over Screen.FontCount asked Qt each time
 * (robot 6.8: a setup program looked for each of its fonts that way) */
static inline QStringList font_families() { static QStringList f = QFontDatabase::families(); return f; }
#endif
#include <string>
#include <set>
#include <vector>
#include <map>
#include <deque>
#include <functional>
#include <algorithm>
#include <cmath>
#include <cstring>
#include <cstdlib>

extern "C" vb_i32 vb_errpend;
extern "C" int vb_lcid;
static const int TPP = 15;                      /* twips per pixel */

/* ---------------------------------------------------------------- strings */
static QString qs(vbs s) { return s.p ? QString::fromUtf16(reinterpret_cast<const char16_t *>(s.p), s.len) : QString(); }
static vbs vs(const QString &q)
{
    vbs r = vb_stmp((vb_i32)q.size());
    if (q.size()) memcpy(r.p, q.utf16(), sizeof(vb_wchar) * (size_t)q.size());
    return r;
}
static vbvar vstr(const QString &q) { return vb_v_s(vs(q)); }
static std::string lower(const char *s) { std::string r(s ? s : ""); for (auto &c : r) c = (char)tolower((unsigned char)c); return r; }
static QString arg_s(int argc, vbvar *argv, int i) { return i < argc ? qs(vb_v2s(argv[i])) : QString(); }
static bool given(int argc, vbvar *argv, int i) { return i < argc && !vb_ismissing(argv[i]) && argv[i].vt != VT_EMPTY; }
static long arg_l(int argc, vbvar *argv, int i, long dflt = 0) { return given(argc, argv, i) ? vb_v2l(argv[i]) : dflt; }
static double arg_d(int argc, vbvar *argv, int i, double dflt = 0) { return given(argc, argv, i) ? vb_v2d(argv[i]) : dflt; }

/* ---------------------------------------------------------------- the application */
static QApplication *app;
static void install_filters();
static bool dark_theme;                         /* the system has a dark theme (ensure_app) */

/* is the theme of the system dark? LBVC_THEME=light|dark first; then Windows: AppsUseLightTheme; Linux:
 * GTK_THEME, MATE / Cinnamon, GNOME's color-scheme / gtk-theme, KDE, and what Qt knows (6.5 and later) */
static bool system_dark()
{
    QByteArray env = qgetenv("LBVC_THEME").toLower();
    if (env == "dark") return true;
    if (env == "light") return false;
#ifdef _WIN32
    QSettings w("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                QSettings::NativeFormat);
    return w.value("AppsUseLightTheme", 1).toInt() == 0;
#else
    /* as the IDE asks (mainwindow.py system_dark): GTK_THEME, the desktop's own setting (MATE:
     * org.mate.interface gtk-theme "Yaru-blue-dark"), GNOME's color-scheme / gtk-theme, KDE; Qt's
     * colorScheme where these say nothing (MATE with Qt's gtk2 platform theme does not tell Qt) */
    auto qt_scheme = []() -> int {
#if QT_VERSION >= QT_VERSION_CHECK(6, 5, 0)
        Qt::ColorScheme cs = QGuiApplication::styleHints()->colorScheme();
        if (cs == Qt::ColorScheme::Dark) return 1;
        if (cs == Qt::ColorScheme::Light) return 0;
#endif
        return -1;
    };
    QByteArray gtk = qgetenv("GTK_THEME").toLower();
    if (!gtk.isEmpty()) return gtk.contains("dark");
    QByteArray desk = qgetenv("XDG_CURRENT_DESKTOP").toLower();
    if (desk.contains("kde")) {
        int q = qt_scheme();
        if (q >= 0) return q == 1;
        QSettings k(QDir::homePath() + "/.config/kdeglobals", QSettings::IniFormat);
        return k.value("General/ColorScheme").toString().contains("dark", Qt::CaseInsensitive);
    }
    auto gs = [](const char *schema, const char *key) -> QByteArray {
        QProcess p;
        p.start("gsettings", QStringList() << "get" << schema << key);
        if (!p.waitForStarted(1000)) return QByteArray();
        if (!p.waitForFinished(1500)) { p.kill(); p.waitForFinished(500); return QByteArray(); }
        return p.exitCode() == 0 ? p.readAllStandardOutput().trimmed().toLower() : QByteArray();
    };
    QByteArray own = desk.contains("mate") ? gs("org.mate.interface", "gtk-theme")
                   : desk.contains("cinnamon") ? gs("org.cinnamon.desktop.interface", "gtk-theme") : QByteArray();
    if (!own.isEmpty()) return own.contains("dark");
    QByteArray cs = gs("org.gnome.desktop.interface", "color-scheme");
    if (cs.contains("prefer-dark")) return true;
    if (cs.contains("prefer-light")) return false;
    QByteArray t = gs("org.gnome.desktop.interface", "gtk-theme");
    if (!t.isEmpty()) return t.contains("dark");
    return qt_scheme() == 1;
#endif
}
static int hwnd_of(QWidget *w);
static void ensure_app()
{
    if (app) return;
    static int argc = 1;
    static char name[] = "vbapp";
    static char *argv[] = { name, nullptr };
    app = new QApplication(argc, argv);
    QApplication::setQuitOnLastWindowClosed(false);
    {                                           /* the icon of a form without one: a small window */
        QIcon ic;
        for (int sz : { 16, 32 }) {
            QPixmap pm(sz, sz);
            pm.fill(Qt::transparent);
            QPainter p(&pm);
            int m = sz / 8, t = sz / 4;
            QRect r(m, m + 1, sz - 2 * m - 1, sz - 2 * m - 2);
            p.fillRect(r, QColor(255, 255, 255));
            p.fillRect(QRect(r.left(), r.top(), r.width(), t), QColor(0, 84, 166));
            p.setPen(QColor(64, 64, 64));
            p.drawRect(r);
            p.end();
            ic.addPixmap(pm);
        }
        QApplication::setWindowIcon(ic);
    }
    /* the colors of the program follow the theme of the system (light: the colors of Windows, as VB6 shows
     * them; dark: a dark set of the same system colors). The Fusion style draws with the palette, so that a
     * control does not mix a dark style with light colors. LBVC_THEME=light|dark chooses it by hand. */
    dark_theme = system_dark();
#ifndef _WIN32
    if (QStyle *st = QStyleFactory::create("Fusion")) QApplication::setStyle(st);
#else
    if (dark_theme) if (QStyle *st = QStyleFactory::create("Fusion")) QApplication::setStyle(st);
#endif
    {
        auto c = [](unsigned rgb) { return QColor((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255); };
        const bool d = dark_theme;
        QPalette pal;
        for (QPalette::ColorGroup g : { QPalette::Active, QPalette::Inactive, QPalette::Disabled }) {
            bool dis = g == QPalette::Disabled;
            unsigned text = d ? (dis ? 0x8A8A8A : 0xEEEEEE) : (dis ? 0x6D6D6D : 0x000000);
            pal.setColor(g, QPalette::Window, c(d ? 0x2D2D2D : 0xF0F0F0));
            pal.setColor(g, QPalette::WindowText, c(text));
            pal.setColor(g, QPalette::Base, c(d ? (dis ? 0x2D2D2D : 0x1E1E1E) : (dis ? 0xF0F0F0 : 0xFFFFFF)));
            pal.setColor(g, QPalette::AlternateBase, c(d ? 0x262626 : 0xF7F7F7));
            pal.setColor(g, QPalette::ToolTipBase, c(d ? 0x383838 : 0xFFFFE1));
            pal.setColor(g, QPalette::ToolTipText, c(d ? 0xEEEEEE : 0x000000));
            pal.setColor(g, QPalette::PlaceholderText, c(d ? 0x8A8A8A : 0x6D6D6D));
            pal.setColor(g, QPalette::Text, c(text));
            pal.setColor(g, QPalette::Button, c(d ? 0x3A3A3A : 0xF0F0F0));
            pal.setColor(g, QPalette::ButtonText, c(text));
            pal.setColor(g, QPalette::BrightText, c(0xFFFFFF));
            pal.setColor(g, QPalette::Light, c(d ? 0x505050 : 0xFFFFFF));
            pal.setColor(g, QPalette::Midlight, c(d ? 0x454545 : 0xE3E3E3));
            pal.setColor(g, QPalette::Mid, c(d ? 0x3A3A3A : 0xA0A0A0));
            pal.setColor(g, QPalette::Dark, c(d ? 0x1E1E1E : 0xA0A0A0));
            pal.setColor(g, QPalette::Shadow, c(d ? 0x0F0F0F : 0x696969));
            pal.setColor(g, QPalette::Highlight, c(d ? (dis ? 0x404040 : 0x3584E4) : (dis ? 0xBFCDDB : 0x3399FF)));
            pal.setColor(g, QPalette::HighlightedText, c(0xFFFFFF));
            pal.setColor(g, QPalette::Link, c(d ? 0x78AEED : 0x0066CC));
            pal.setColor(g, QPalette::LinkVisited, c(d ? 0xC061CB : 0x800080));
        }
        QApplication::setPalette(pal);
    }
    QFont f = QApplication::font();
    f.setPixelSize(11);                         /* MS Sans Serif 8.25 */
    QApplication::setFont(f);
    install_filters();
}

/* things VB6 does from the message loop (GotFocus/LostFocus ...): run in order on idle,
 * or earlier by DoEvents and SendKeys */
static std::deque<std::function<void()>> pending;
static bool flushing, flush_scheduled;
static void flush_pending()
{
    flush_scheduled = false;
    if (flushing) return;
    flushing = true;
    while (!pending.empty() && !vb_errpend) { auto f = pending.front(); pending.pop_front(); f(); }
    flushing = false;
}
static void post(std::function<void()> f)
{
    pending.push_back(std::move(f));
    if (!flush_scheduled) { flush_scheduled = true; QTimer::singleShot(0, flush_pending); }
}

/* ---------------------------------------------------------------- colors */
/* the system colors (&H800000nn&, COLOR_nn) with a dark theme of the system: the same roles, dark */
static const unsigned sys_dark[] = {
    0x3C3C3C, 0x000000, 0x303030, 0x2A2A2A, 0x2D2D2D, 0x1E1E1E, 0x5A5A5A, 0xEEEEEE, 0xEEEEEE, 0xFFFFFF,
    0x3A3A3A, 0x3A3A3A, 0x1A1A1A, 0x3584E4, 0xFFFFFF, 0x2D2D2D, 0x1E1E1E, 0x8A8A8A, 0xEEEEEE, 0xAAAAAA,
    0x505050, 0x0F0F0F, 0x454545, 0xEEEEEE, 0x383838, 0x000000, 0x78AEED, 0x303030, 0x2A2A2A, 0x3584E4, 0x2D2D2D };
static long norm_color(double d) { return (long)(vb_i32)(vb_u32)(long long)d; }
static QColor qcolor_in(long c, bool dark);
static QColor qcolor(long c) { return qcolor_in(c, dark_theme); }
/* the text color of a control: with a dark theme of the system, a system color (ButtonText ...) on a light
 * BackColor of the program's own (&H00FFFFFF&) is the one of the light theme: the label of the Data Form
 * Wizard's record bar was white on white (DEV-9 df3) */
static QColor fore_on(long fore, long back)
{
    if (dark_theme && (fore & 0x80000000L) && !(back & 0x80000000L) && qcolor(back).lightness() >= 128)
        return qcolor_in(fore, false);
    return qcolor(fore);
}
static QColor qcolor_in(long c, bool dark)
{
    if (c & 0x80000000L) {                      /* system colors of Windows 7 (dark: sys_dark) */
        static const unsigned sys[] = {
            0xC8C8C8, 0x000000, 0x99B4D1, 0xBFCDDB, 0xF0F0F0, 0xFFFFFF, 0x646464, 0x000000, 0x000000, 0x000000,
            0xB4B4B4, 0xF4F7FC, 0xABABAB, 0x3399FF, 0xFFFFFF, 0xF0F0F0, 0xA0A0A0, 0x6D6D6D, 0x000000, 0x434E54,
            0xFFFFFF, 0x696969, 0xE3E3E3, 0x000000, 0xFFFFE1, 0x000000, 0x0066CC, 0xB9D1EA, 0xD7E4F2, 0xF0F0F0, 0x000000 };
        unsigned i = (unsigned)(c & 0xFF);
        const unsigned *t = dark ? sys_dark : sys;
        unsigned rgb = i < sizeof sys / sizeof sys[0] ? t[i] : 0;
        return QColor((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
    }
    return QColor(c & 255, (c >> 8) & 255, (c >> 16) & 255);
}
static long vbcolor(QRgb p) { return (long)(qRed(p) | (qGreen(p) << 8) | (qBlue(p) << 16)); }

/* ---------------------------------------------------------------- fonts */
struct VFont {
    QString name = "MS Sans Serif";
    double size = 8.25;
    bool bold = false, italic = false, underline = false, strike = false;
    int weight = 400, charset = 0;
};
/* Windows makes fonts in whole pixels: 10 points -> 13 px -> 9.75 points */
static double real_size(double pt) { double px = std::floor(pt * 4.0 / 3.0 + 0.5); if (px < 1) px = 1; return px * 0.75; }
static vb_i32 font_handle(const VFont &f);   /* winapi_gui.inc */
static QFont qfont(const VFont &f, double scale = 1.0)
{
    QFont q(f.name);
    q.setPixelSize((int)std::lround(f.size * 4.0 / 3.0 * scale));
    q.setBold(f.bold); q.setItalic(f.italic); q.setUnderline(f.underline); q.setStrikeOut(f.strike);
    return q;
}
static void font_from_props(VFont *f, const char *const *props);

/* ---------------------------------------------------------------- graphics (forms, PictureBox) */
struct Gfx {
    QImage img;
    QWidget *w = nullptr;
    int pw = 0, ph = 0;                         /* the client area in pixels */
    long back = 0x8000000F, fore = 0x80000012, fillcolor = 0;
    int drawwidth = 1, drawstyle = 0, drawmode = 13, fillstyle = 1, scalemode = 1;
    int palettemode = 0;                        /* the colors are true colors: the palette changes nothing */
    double sl = 0, st = 0, sw = 0, sh = 0;      /* ScaleLeft, ScaleTop, ScaleWidth, ScaleHeight */
    double cx = 0, cy = 0;                      /* CurrentX, CurrentY */
    bool autoredraw = false, fonttransparent = true, clipcontrols = true;
    VFont *font = nullptr;
    QImage pic;                                 /* the Picture: the background that Cls restores */
    double tpp = 15;                            /* twips per pixel of img (the Printer: 7.5 = 192 dpi) */
};

static double units_per_twip(int mode, bool vert)
{
    switch (mode) {
    case 2: return 1.0 / 20; case 3: return 1.0 / TPP; case 4: return vert ? 1.0 / 240 : 1.0 / 120;
    case 5: return 1.0 / 1440; case 6: return 1.0 / 56.6929133858; case 7: return 1.0 / 566.929133858;
    default: return 1.0;
    }
}
static void gfx_rescale(Gfx *g)
{
    if (g->scalemode == 0) return;
    g->sl = g->st = 0;
    g->sw = g->pw * g->tpp * units_per_twip(g->scalemode, false);
    g->sh = g->ph * g->tpp * units_per_twip(g->scalemode, true);
}
static void gfx_resize(Gfx *g, int pw, int ph)
{
    if (pw < 0) pw = 0;
    if (ph < 0) ph = 0;
    int ow = g->img.width(), oh = g->img.height();
    g->pw = pw; g->ph = ph;
    if (pw > ow || ph > oh || g->img.isNull()) {
        QImage n(std::max(pw, ow), std::max(ph, oh), QImage::Format_RGB32);
        n.fill(qcolor(g->back));
        if (!g->pic.isNull()) { QPainter p(&n); p.drawImage(0, 0, g->pic); }
        if (!g->img.isNull()) { QPainter p(&n); p.drawImage(0, 0, g->img); }
        g->img = n;
    }
    gfx_rescale(g);
}
static double px_x(Gfx *g, double x) { return g->sw ? (x - g->sl) * g->pw / g->sw : 0; }
static double px_y(Gfx *g, double y) { return g->sh ? (y - g->st) * g->ph / g->sh : 0; }
static double un_w(Gfx *g, double px) { return g->pw ? px * g->sw / g->pw : 0; }
static double un_h(Gfx *g, double px) { return g->ph ? px * g->sh / g->ph : 0; }
static int ipx(double v) { return (int)std::floor(v + 0.5); }
static bool painting;                            /* in a Paint event: drawing does not ask for another one */
static void gfx_update(Gfx *g) { if (g->w && !painting) g->w->update(); }
static void gfx_clear(Gfx *g)
{
    g->img.fill(qcolor(g->back));
    if (!g->pic.isNull()) { QPainter p(&g->img); p.drawImage(0, 0, g->pic); }
    gfx_update(g);
}
static void gfx_cls(Gfx *g) { gfx_clear(g); g->cx = g->cy = 0; }

static void put_px(Gfx *g, int x, int y, const QColor &c)
{
    if (x >= 0 && y >= 0 && x < g->img.width() && y < g->img.height()) g->img.setPixelColor(x, y, c);
}
/* a line like GDI LineTo: the last point is not drawn */
static void gfx_line(Gfx *g, int x0, int y0, int x1, int y1, const QColor &c)
{
    if (g->drawwidth > 1) {
        QPainter p(&g->img);
        p.setPen(QPen(c, g->drawwidth, Qt::SolidLine, Qt::RoundCap));
        p.drawLine(x0, y0, x1, y1);
        return;
    }
    int dx = std::abs(x1 - x0), sx = x0 < x1 ? 1 : -1, dy = -std::abs(y1 - y0), sy = y0 < y1 ? 1 : -1, err = dx + dy;
    while (x0 != x1 || y0 != y1) {
        put_px(g, x0, y0, c);
        int e2 = 2 * err;
        if (e2 >= dy) { err += dy; x0 += sx; }
        if (e2 <= dx) { err += dx; y0 += sy; }
    }
}

static long color_arg(Gfx *g, int argc, vbvar *argv, int i) { return given(argc, argv, i) ? vb_v2l(argv[i]) : g->fore; }

static vbvar gfx_method(Gfx *g, const std::string &n, int argc, vbvar *argv)
{
    if (n == "$pset") {                         /* step, x, y, color */
        double x = vb_v2d(argv[1]), y = vb_v2d(argv[2]);
        if (vb_v2bool(argv[0])) { x += g->cx; y += g->cy; }
        QColor c = qcolor(color_arg(g, argc, argv, 3));
        if (g->drawwidth > 1) {
            QPainter p(&g->img); p.setPen(Qt::NoPen); p.setBrush(c);
            p.drawEllipse(QPointF(px_x(g, x), px_y(g, y)), g->drawwidth / 2.0, g->drawwidth / 2.0);
        } else put_px(g, ipx(px_x(g, x)), ipx(px_y(g, y)), c);
        g->cx = x; g->cy = y;
    } else if (n == "$line") {                  /* has p1, step1, x1, y1, step2, x2, y2, color, box */
        double x1 = g->cx, y1 = g->cy;
        if (vb_v2bool(argv[0])) {
            x1 = vb_v2d(argv[2]); y1 = vb_v2d(argv[3]);
            if (vb_v2bool(argv[1])) { x1 += g->cx; y1 += g->cy; }
        }
        double x2 = vb_v2d(argv[5]), y2 = vb_v2d(argv[6]);
        if (vb_v2bool(argv[4])) { x2 += x1; y2 += y1; }
        QColor c = qcolor(color_arg(g, argc, argv, 7));
        int box = (int)arg_l(argc, argv, 8);
        int a = ipx(px_x(g, x1)), b = ipx(px_y(g, y1)), cc = ipx(px_x(g, x2)), d = ipx(px_y(g, y2));
        if (box == 2) {
            QPainter p(&g->img);
            p.fillRect(QRect(QPoint(std::min(a, cc), std::min(b, d)), QPoint(std::max(a, cc), std::max(b, d))), c);
        } else if (box == 1) {
            QRect r(QPoint(std::min(a, cc), std::min(b, d)), QPoint(std::max(a, cc), std::max(b, d)));
            QPainter p(&g->img);
            if (g->fillstyle == 0) p.fillRect(r, qcolor(g->fillcolor));
            p.setPen(QPen(c, g->drawwidth));
            p.drawRect(r.adjusted(0, 0, -1 + (g->drawwidth > 1 ? 0 : 1) - 1, -1 + (g->drawwidth > 1 ? 0 : 1) - 1));
        } else gfx_line(g, a, b, cc, d, c);
        g->cx = x2; g->cy = y2;
    } else if (n == "$circle") {                /* step, x, y, r, color, start, end, aspect */
        double x = vb_v2d(argv[1]), y = vb_v2d(argv[2]), r = vb_v2d(argv[3]);
        if (vb_v2bool(argv[0])) { x += g->cx; y += g->cy; }
        QColor c = qcolor(color_arg(g, argc, argv, 4));
        double aspect = arg_d(argc, argv, 7, 1.0);
        double rx = std::fabs(px_x(g, g->sl + r) - px_x(g, g->sl)), ry = rx;
        if (aspect > 1) { ry = rx; rx = ry / aspect; } else ry = rx * aspect;
        QPainter p(&g->img);
        p.setPen(QPen(c, g->drawwidth));
        if (g->fillstyle == 0) p.setBrush(qcolor(g->fillcolor));
        QRectF rect(px_x(g, x) - rx, px_y(g, y) - ry, 2 * rx, 2 * ry);
        if (given(argc, argv, 5) || given(argc, argv, 6)) {
            double s = std::fabs(arg_d(argc, argv, 5, 0)), e = std::fabs(arg_d(argc, argv, 6, 2 * M_PI));
            double span = e - s; if (span < 0) span += 2 * M_PI;
            p.drawArc(rect, (int)std::lround(s * 180 / M_PI * 16), (int)std::lround(span * 180 / M_PI * 16));
        } else p.drawEllipse(rect);
        g->cx = x; g->cy = y;
    } else if (n == "$scale") {
        if (argc < 4) { g->scalemode = 1; gfx_rescale(g); }
        else {
            double x1 = vb_v2d(argv[0]), y1 = vb_v2d(argv[1]), x2 = vb_v2d(argv[2]), y2 = vb_v2d(argv[3]);
            g->scalemode = 0; g->sl = x1; g->st = y1; g->sw = x2 - x1; g->sh = y2 - y1;
        }
        return vb_v_empty();
    } else if (n == "$print") {                 /* text, newline */
        QString text = arg_s(argc, argv, 0);
        bool nl = vb_v2bool(argv[1]) != 0;
        QFont f = g->font ? qfont(*g->font, TPP / g->tpp) : QApplication::font();
        QFontMetrics fm(f);
        QPainter p(&g->img);
        p.setFont(f); p.setPen(qcolor(g->fore));
        double zone = fm.averageCharWidth() * 14.0;
        QStringList lines = text.split('\n');
        for (int li = 0; li < lines.size(); li++) {
            QStringList parts = lines[li].split('\t');
            for (int k = 0; k < parts.size(); k++) {
                if (k) {                        /* a comma: the next print zone */
                    double x = px_x(g, g->cx);
                    double nx = (std::floor(x / zone) + 1) * zone;
                    g->cx = g->sl + un_w(g, nx);
                }
                double x = px_x(g, g->cx), y = px_y(g, g->cy);
                if (!g->fonttransparent) p.fillRect(QRectF(x, y, fm.horizontalAdvance(parts[k]), fm.height()), qcolor(g->back));
                p.drawText(QPointF(x, y + fm.ascent()), parts[k]);
                g->cx += un_w(g, fm.horizontalAdvance(parts[k]));
            }
            if (li < lines.size() - 1 || nl) {
                if (li == lines.size() - 1 && !nl) break;
                g->cx = g->sl;
                g->cy += un_h(g, fm.height());
            }
        }
    } else return vb_v_empty();
    gfx_update(g);
    return vb_v_empty();
}

/* members shared by forms and PictureBoxes; *found = 0 when n is not one */
static const QImage *image_of_value(vbvar v);
static vbobj new_pic(const QImage &img, int type);
static vbvar gfx_member(Gfx *g, const std::string &n, bool set, vbvar val, int argc, vbvar *argv, int *found)
{
    *found = 1;
    if (n == "paintpicture") {                  /* Picture, X1, Y1, [Width1, Height1, X2, Y2, Width2, Height2, Opcode] */
        const QImage *src = argc > 0 ? image_of_value(argv[0]) : nullptr;
        if (!src) { vb_raise(argc > 0 ? 481 : 449); return vb_v_empty(); }
        if (src->isNull()) return vb_v_empty();
        auto pw = [&](double w) { return g->sw ? w * g->pw / g->sw : 0; };
        auto ph = [&](double h) { return g->sh ? h * g->ph / g->sh : 0; };
        double x1 = px_x(g, arg_d(argc, argv, 1)), y1 = px_y(g, arg_d(argc, argv, 2));
        double w1 = given(argc, argv, 3) ? pw(arg_d(argc, argv, 3)) : src->width();
        double h1 = given(argc, argv, 4) ? ph(arg_d(argc, argv, 4)) : src->height();
        double x2 = given(argc, argv, 5) ? pw(arg_d(argc, argv, 5)) : 0, y2 = given(argc, argv, 6) ? ph(arg_d(argc, argv, 6)) : 0;
        double w2 = given(argc, argv, 7) ? pw(arg_d(argc, argv, 7)) : src->width() - x2;
        double h2 = given(argc, argv, 8) ? ph(arg_d(argc, argv, 8)) : src->height() - y2;
        long op = arg_l(argc, argv, 9, 0xCC0020);
        QImage part = src->copy(QRectF(std::min(x2, x2 + w2), std::min(y2, y2 + h2), std::abs(w2), std::abs(h2)).toRect());
        if ((w1 < 0) != (w2 < 0) || (h1 < 0) != (h2 < 0))      /* a negative width or height: a mirror */
            part = part.mirrored((w1 < 0) != (w2 < 0), (h1 < 0) != (h2 < 0));
        QRectF dst(std::min(x1, x1 + w1), std::min(y1, y1 + h1), std::abs(w1), std::abs(h1));
        QPainter p(&g->img);
        p.setRenderHint(QPainter::SmoothPixmapTransform, false);
        switch (op) {
        case 0x8800C6: p.setCompositionMode(QPainter::RasterOp_SourceAndDestination); break;   /* vbSrcAnd */
        case 0xEE0086: p.setCompositionMode(QPainter::RasterOp_SourceOrDestination); break;    /* vbSrcPaint */
        case 0x660046: p.setCompositionMode(QPainter::RasterOp_SourceXorDestination); break;   /* vbSrcInvert */
        case 0x330008: p.setCompositionMode(QPainter::RasterOp_NotSource); break;              /* vbNotSrcCopy */
        default: break;
        }
        p.drawImage(dst, part);
        p.end();
        gfx_update(g);
        return vb_v_empty();
    }
    if (n[0] == '$') return gfx_method(g, n, argc, argv);
    if (n == "cls") { gfx_cls(g); return vb_v_empty(); }
    if (n == "point") {
        int x = ipx(px_x(g, arg_d(argc, argv, 0))), y = ipx(px_y(g, arg_d(argc, argv, 1)));
        if (x < 0 || y < 0 || x >= g->pw || y >= g->ph) return vb_v_l(-1);
        return vb_v_l((vb_i32)vbcolor(g->img.pixel(x, y)));
    }
    if (n == "textwidth" || n == "textheight") {
        QFontMetrics fm(g->font ? qfont(*g->font, TPP / g->tpp) : QApplication::font());
        QStringList lines = arg_s(argc, argv, 0).split(QStringLiteral("\r\n"));
        if (n == "textheight") return vb_v_f((float)un_h(g, fm.height() * lines.size()));
        int w = 0; for (auto &l : lines) w = std::max(w, (int)fm.horizontalAdvance(l));
        return vb_v_f((float)un_w(g, w));
    }
    if (n == "currentx") { if (set) { g->cx = vb_v2d(val); return vb_v_empty(); } return vb_v_f((float)g->cx); }
    if (n == "currenty") { if (set) { g->cy = vb_v2d(val); return vb_v_empty(); } return vb_v_f((float)g->cy); }
    if (n == "scalemode") {
        if (!set) return vb_v_i((vb_i16)g->scalemode);
        int m = vb_v2i(val);
        if (m < 0 || m > 7) { vb_raise(380); return vb_v_empty(); }
        g->scalemode = m; gfx_rescale(g);
        return vb_v_empty();
    }
    if (n == "scaleleft" || n == "scaletop" || n == "scalewidth" || n == "scaleheight") {
        double *p = n == "scaleleft" ? &g->sl : n == "scaletop" ? &g->st : n == "scalewidth" ? &g->sw : &g->sh;
        if (!set) return vb_v_f((float)*p);
        double v = vb_v2d(val);
        if ((n == "scalewidth" || n == "scaleheight") && v == 0) { vb_raise(380); return vb_v_empty(); }
        *p = v; g->scalemode = 0;
        return vb_v_empty();
    }
    if (n == "scalex" || n == "scaley") {      /* value, from mode, to mode */
        bool vert = n == "scaley";
        double v = arg_d(argc, argv, 0);
        int from = (int)arg_l(argc, argv, 1, 8), to = (int)arg_l(argc, argv, 2, g->scalemode);
        auto twips = [&](int m, double x) { return m == 0 ? (vert ? x * g->ph * g->tpp / g->sh : x * g->pw * g->tpp / g->sw) : m == 8 ? x * 1440 / 2540 : x / units_per_twip(m, vert); };
        double t = twips(from, v);
        double r = to == 0 ? (vert ? t * g->sh / (g->ph * g->tpp) : t * g->sw / (g->pw * g->tpp)) : to == 8 ? t * 2540 / 1440 : t * units_per_twip(to, vert);
        return vb_v_f((float)r);
    }
    if (n == "drawwidth") { if (set) { int v = vb_v2i(val); if (v < 1) { vb_raise(380); return vb_v_empty(); } g->drawwidth = v; return vb_v_empty(); } return vb_v_i((vb_i16)g->drawwidth); }
    if (n == "drawstyle") { if (set) { g->drawstyle = vb_v2i(val); return vb_v_empty(); } return vb_v_i((vb_i16)g->drawstyle); }
    if (n == "drawmode") { if (set) { g->drawmode = vb_v2i(val); return vb_v_empty(); } return vb_v_i((vb_i16)g->drawmode); }
    if (n == "fillstyle") { if (set) { g->fillstyle = vb_v2i(val); return vb_v_empty(); } return vb_v_i((vb_i16)g->fillstyle); }
    if (n == "fillcolor") { if (set) { g->fillcolor = vb_v2l(val); return vb_v_empty(); } return vb_v_l((vb_i32)g->fillcolor); }
    if (n == "autoredraw") { if (set) { g->autoredraw = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(g->autoredraw ? -1 : 0); }
    if (n == "palettemode") { if (set) { int v = vb_v2i(val); if (v < 0 || v > 2) { vb_raise(380); return vb_v_empty(); } g->palettemode = v; return vb_v_empty(); } return vb_v_i((vb_i16)g->palettemode); }
    if (n == "palette") return vb_v_empty();
    if (n == "image") {                         /* the picture AutoRedraw keeps, as a Picture */
        if (set) { vb_raise(383); return vb_v_empty(); }
        QImage img = g->img.copy(0, 0, std::max(1, g->pw), std::max(1, g->ph));
        return vb_vtmp(vb_v_obj(new_pic(img, 1)));
    }
    if (n == "clipcontrols") { if (set) { g->clipcontrols = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(g->clipcontrols ? -1 : 0); }
    if (n == "hdc") return vb_v_empty();       /* (answered by the caller) */
    if (n == "fonttransparent") { if (set) { g->fonttransparent = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(g->fonttransparent ? -1 : 0); }
    *found = 0;
    return vb_v_empty();
}

static void gfx_from_props(Gfx *g, const char *const *props, double twips_w, double twips_h);

/* the drawing surface: the client area of a form, a PictureBox */
class Canvas : public QFrame {
public:
    Gfx *g;
    std::function<void(int, int)> on_resize;
    std::function<void()> on_paint;             /* the Paint event (AutoRedraw = False) */
    bool transparent = false;                   /* a user control with BackStyle = 0: the form shows through */
    Canvas(Gfx *gfx, QWidget *parent) : QFrame(parent), g(gfx) { g->w = this; setAttribute(Qt::WA_OpaquePaintEvent); }
    void set_transparent() { transparent = true; setAttribute(Qt::WA_OpaquePaintEvent, false); setAutoFillBackground(false); }
protected:
    void paintEvent(QPaintEvent *) override
    {
        if (on_paint && !g->autoredraw && !painting) { painting = true; on_paint(); painting = false; }
        QPainter p(this);
        QRect cr = contentsRect();
        if (!transparent) {
            p.fillRect(cr, qcolor(g->back));
            p.drawImage(cr.topLeft(), g->img, QRect(0, 0, cr.width(), cr.height()));
        }
        drawFrame(&p);
    }
    void resizeEvent(QResizeEvent *e) override
    {
        QFrame::resizeEvent(e);
        QRect cr = contentsRect();
        gfx_resize(g, cr.width(), cr.height());
        if (on_resize) on_resize(cr.width(), cr.height());
    }
};

/* ---------------------------------------------------------------- state */
enum Kind { K_COMMAND, K_LABEL, K_TEXT, K_CHECK, K_OPTION, K_LIST, K_COMBO, K_TIMER, K_FRAME, K_PICTURE, K_MENU,
            K_IMAGE, K_SHAPE, K_LINE, K_HSCROLL, K_VSCROLL, K_DRIVE, K_DIR, K_FILE, K_DIALOG, K_OCX, K_USER, K_OTHER,
            K_DATA };

struct FormState;
struct Ocx;                                     /* the ActiveX controls (ocx.inc) */
struct DataCtl;                                 /* the Data control (data.inc) */
struct OcxType;
struct Ctl;
static const OcxType *ocx_type(const std::string &t);
static Ocx *ocx_new(Ctl *c);
static QWidget *ocx_make(Ctl *c, QWidget *parent);
static void ocx_load(Ctl *c, const char *const *p);
static vbvar ocx_member(Ctl *c, const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found);
static bool ocx_clicks(Ocx *x);
static bool ocx_inside(Ocx *x, QWidget *w, QPoint lp);
static bool ocx_eat_mouse(Ocx *x, QMouseEvent *me);
static void ocx_resized(Ctl *c);
static void ocx_shown(FormState *st);
static void ocx_loaded(FormState *st);
static void ocx_delete(Ocx *x);
static const char *ocx_default(const std::string &type);
/* the CommonDialog control (MSComDlg.CommonDialog, comdlg32.ocx) */
struct CDlg {
    QString filename, filetitle, filter, title, initdir, defaultext, fontname = "MS Sans Serif";
    int filterindex = 0, copies = 1, frompage = 0, topage = 0, min = 0, max = 0, orientation = 1;
    long flags = 0, color = 0;
    bool cancelerror = false, bold = false, italic = false, underline = false, strike = false;
    double fontsize = 8.25;
    QString helpfile, helpkey;                  /* ShowHelp (help.inc) */
    vb_i32 helpcommand = 0, helpcontext = 0;
};
struct Ctl {
    Kind kind;
    std::string type;                           /* CommandButton ... */
    QString name;
    int index = -1;                             /* in a control array */
    FormState *form = nullptr;
    vbobj self = nullptr;
    vbobj container = nullptr;                  /* Frame / PictureBox, or none (the form) */
    QPointer<QWidget> w;
    QTimer *timer = nullptr;
    double left = 0, top = 0, width = 0, height = 0;
    bool visible = true, enabled = true, checked = false;
    int tabindex = 0;
    QString tag, caption, shortcut, tooltip;
    std::map<std::string, vb_i32> misc;         /* HelpContextID, OLEDragMode ...: kept, nothing more */
    vbobj btnpic[3] = { nullptr, nullptr, nullptr };   /* Picture, DownPicture, DisabledPicture of a button */
    int value = 0;                              /* CheckBox */
    int interval = 0;
    std::vector<long> itemdata;
    int newindex = -1;
    long back = 0x8000000F, fore = 0x80000012;
    VFont font;
    Gfx *g = nullptr;                           /* PictureBox */
    QAction *act = nullptr;                     /* Menu */
    QMenu *menu = nullptr;
    vbobj picture = nullptr;                    /* Image, PictureBox: a Picture object */
    bool stretch = false, autosize = false, look_done = false, wordwrap = false;
    int alignment = 0;                          /* Label: 0 left, 1 right, 2 center */
    int align = 0;                              /* Align: 1 top, 2 bottom, 3 left, 4 right (PictureBox ...) */
    bool align_keyed = false; double align_key = 0;   /* the place in the .frm: the order of the aligned controls */
    int shape = 0, backstyle = 0, borderstyle = 1, borderwidth = 1, fillstyle = 1;   /* Shape, Line */
    long bordercolor = 0x80000008, fillcolor = 0;
    double x1 = 0, y1 = 0, x2 = 0, y2 = 0;      /* Line */
    int smin = 0, smax = 32767, sval = 0, small = 1, large = 1;   /* scroll bars */
    QString path, pattern = "*.*";              /* DirListBox, FileListBox, DriveListBox (drive) */
    bool f_archive = true, f_hidden = false, f_normal = true, f_readonly = true, f_system = false;
    bool dbl = false;                           /* a double click: no Click on its MouseUp */
    bool causesval = true, added = false;       /* CausesValidation; made by Controls.Add */
    int dragmode = 0, mousepointer = 0;
    vbobj mouseicon = nullptr;                  /* MousePointer 99: the cursor */
    const vbf_ctldef *def = nullptr;
    CDlg *dlg = nullptr;                        /* CommonDialog */
    Ocx *ocx = nullptr;                         /* a control of an .ocx */
    vbobj uc = nullptr;                         /* a user control (.ctl): its object */
    vbobj bag = nullptr;                        /* Load of an element: the PropertyBag of its ReadProperties */
    bool uc_shown = false;                      /* its Show came (Hide when it goes) */
    bool uc_sized = false;                      /* its first Resize came */
    bool uc_loading = false, uc_ready = false;  /* in its ReadProperties / InitProperties; after them */
    DataCtl *data = nullptr;                    /* a Data control */
    QString datafield, datasource;              /* a bound control: DataField, DataSource (a Data control's name) */
    QString datamember;                         /* DataMember: of a data source class (Set DataSource = obj) */
    bool datachanged = false;
    vbobj dsobj = nullptr;                      /* Set DataSource = an object (a Recordset, a class): counted */
    vbobj dsrs = nullptr;                       /* the ADO Recordset it listens to (counted) */
    vbobj dfmt = nullptr;                       /* DataFormat: a StdDataFormat (counted) */
    bool dfmt_on = false;                       /* from the .frm or Set in code: the ADO binding uses it */
};

struct CtlArray { QString name; std::map<int, vbobj> elems; FormState *form = nullptr; };

struct FormState {
    vbobj form = nullptr;
    const vbf_formdef *def = nullptr;
    QPointer<QWidget> w;
    QPointer<Canvas> canvas;
    QMenuBar *menubar = nullptr;
    bool loaded = false, visible = false, unloading = false, shown_once = false;
    bool activate_pending = false;       /* Show: its Activate is posted (activate_now) */
    bool sizing = false;                        /* the runtime sizes the window: its canvas has the size it set */
    bool in_unload = false;                     /* QueryUnload / Unload are running: Unload Me there does nothing */
    std::vector<vbobj> ctls;                    /* every control object (load order) */
    std::map<std::string, vbobj> byname;        /* lower name -> control, or control array object */
    QString caption, tag;
    double left = 0, top = 0, clientw = 4800, clienth = 3600;
    bool enabled = true, keypreview = false;
    bool whatsthis = false, whatsthisbtn = false;   /* WhatsThisHelp, WhatsThisButton (help.inc) */
    vb_i32 helpctx = 0;                         /* HelpContextID: F1 (help.inc) */
    vbobj focus = nullptr;                      /* ActiveControl (weak: one of ctls) */
    vbobj picture = nullptr;
    bool dbl = false;
    Gfx g;
    VFont font;
    /* the window */
    int bstyle = 2, startpos = 3, wstate = 0;
    bool controlbox = true, minbtn = true, maxbtn = true, taskbar = true, moveable = true;
    double norm_cw = 0, norm_ch = 0;            /* the client size before Minimize/Maximize */
    double winw = 0, winh = 0;                  /* Width, Height as VB6 reports them (twips) */
    /* MDI */
    bool mdi = false, mdichild = false, autoshow = true;
    QPointer<QMdiArea> area;                    /* MDIForm */
    QPointer<QMdiSubWindow> sub;                /* MDI child */
    vbobj active_child = nullptr;               /* MDIForm: ActiveForm (weak) */
    vbobj icon = nullptr;
    int mousepointer = 0;
    vbobj mouseicon = nullptr;
    /* a user control: the control of the form that has it (its Extender) */
    Ctl *host = nullptr;
    std::map<std::string, vbvar> extra;         /* UserControl properties kept as they are set (MaskPicture ...) */
};

static std::vector<vbobj> loaded_forms;          /* the Forms collection (holds a reference each) */
static vbobj active_form;
static bool running;

static void unload_form(vbobj o, int mode);
static void form_show(vbobj o, int modal);
static void activate_now(vbobj o);
static void monkey_start();
/* user controls (uctl.inc) */
static const vb_class *uctl_class(const char *cls);
static void uctl_load(vbobj o);
static QWidget *uctl_make(Ctl *c, QWidget *parent);
static void uctl_after(Ctl *c);
static vbvar uctl_member(Ctl *c, const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found);
static vbvar uctl_self_member(vbobj o, const std::string &n, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found);
static void uctl_geometry(Ctl *c);
static void uctl_visible(Ctl *c);
static void uctl_shown(FormState *st);
static void uctl_release(Ctl *c);
static vbobj uctl_copy_bag(Ctl *src);
static bool uctl_can_focus(Ctl *c);
static QWidget *uctl_focus_widget(Ctl *c);
static Ctl *uctl_inner_focus(Ctl *c);
static bool uctl_access_key(FormState *st, int ascii);
static vbobj new_ctl_object(FormState *st, const vbf_ctldef *d, Kind k, const std::string &type);
/* the Data control (data.inc) */
static void data_load(Ctl *c, const char *const *p);
static QWidget *data_make(Ctl *c, QWidget *parent);
static vbvar data_member(Ctl *c, const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found);
static vbvar bound_member(Ctl *c, const std::string &n, bool set, vbvar val, int *found);
static void data_after_load(FormState *st);
static void adodc_after_load(FormState *st);    /* data.inc: the ADO Data controls open before Form_Load */
static void bound_release(Ctl *c);              /* data.inc: a bound control stops listening */
static void dfmt_from_props(Ctl *c, const char *const *p);   /* data.inc: the DataFormat of the .frm */
static bool ocx_data_aware(Ocx *x);             /* ocx.inc: DataSource ... are the control's own */
static bool ocx_focusable(Ocx *x);              /* ocx.inc: an .ocx control that takes the focus (MaskEdBox) */
static void data_unload(Ctl *c);

static const vbf_formdef *formdef(vbobj o) { return o ? (const vbf_formdef *)o->cls->form : nullptr; }
static FormState *fstate(vbobj o)
{
    o = vb_oid(o);
    if (!o->ext) { FormState *st = new FormState; st->form = o; st->def = formdef(o); o->ext = st; }
    return (FormState *)o->ext;
}

/* ---------------------------------------------------------------- events */
static bool fire(vbobj form, const std::string &name, int argc, vbvar *argv)
{
    const vbf_formdef *d = formdef(form);
    if (!d || vb_errpend) return false;
    for (int i = 0; i < d->nevents; i++)
        if (name == d->events[i].name) {
            vb_oaddref(form);
            d->events[i].fn(form, argc, argv);
            if (vb_errpend) vb_unhandled();     /* an error in an event procedure ends the program */
            vb_orelease(form);
            return true;
        }
    return false;
}

/* WithEvents mText As TextBox, WithEvents f As Form: the handler table is event name -> procedure */
static bool fire_sinks(vbobj self, const char *event, int argc, vbvar *argv)
{
    bool r = false;
    vbobj *ls; const vb_fn **ts;
    vb_i32 ns = vb_sinks_begin(self, &ls, &ts);
    for (vb_i32 k = 0; k < ns && !vb_errpend; k++)
        for (const vb_member *mm = (const vb_member *)ts[k]; mm->name; mm++)
            if (!strcmp(mm->name, event)) {
                mm->fn(ls[k], argc, argv);
                if (vb_errpend) vb_unhandled();
                r = true;
                break;
            }
    vb_sinks_end(ls, ts, ns);
    return r;
}
static bool data_filling;
static bool fire_ctl(Ctl *c, const char *event, int argc = 0, vbvar *argv = nullptr)
{
    if (!c->form) return false;
    /* a bound control the program or the user changed (not the Data control filling it) */
    if (!data_filling && !c->datafield.isEmpty() && (!strcmp(event, "change") || (c->kind == K_CHECK && !strcmp(event, "click"))))
        c->datachanged = true;
    /* a control a BindingCollection binds: its Text / Value changed (PropertyChanged; VB6 oracle f_dfmt) */
    if (vb_dbhooks.prop_changed && (!strcmp(event, "change") || (c->kind == K_CHECK && !strcmp(event, "click"))) &&
        (c->kind == K_TEXT || c->kind == K_COMBO || c->kind == K_CHECK))
        vb_dbhooks.prop_changed(c->self, c->kind == K_CHECK ? "value" : "text");
    std::string n = lower(c->name.toUtf8().constData()) + "_" + event;
    std::vector<vbvar> a;
    if (c->index >= 0) a.push_back(vb_v_i((vb_i16)c->index));
    for (int i = 0; i < argc; i++) a.push_back(argv[i]);
    vb_i32 m = vb_mark();
    vbobj self = c->self;
    vb_oaddref(self);
    bool r = fire(c->form->form, n, (int)a.size(), a.data());
    for (int i = 0; i < argc; i++) argv[i] = a[i + (c->index >= 0 ? 1 : 0)];
    if (fire_sinks(self, event, argc, argv)) r = true;
    vb_orelease(self);
    vb_release(m);
    return r;
}

static std::string ev_prefix(vbobj form) { const vbf_formdef *d = formdef(form); return d && d->mdi == 2 ? "usercontrol_" : d && d->mdi ? "mdiform_" : "form_"; }
static bool fire_form(vbobj form, const char *event, int argc = 0, vbvar *argv = nullptr)
{
    vb_i32 m = vb_mark();
    bool r = fire(form, ev_prefix(form) + event, argc, argv);
    vbobj self = vb_oid(form);
    if (self && self->sinks) {
        vb_oaddref(self);
        if (fire_sinks(self, event, argc, argv)) r = true;
        vb_orelease(self);
    }
    vb_release(m);
    return r;
}

/* ---------------------------------------------------------------- the form window */
/* the menu bar of a form: 20 pixels high, as in Windows 7, whatever the style (the frame of VB6 counts it so:
 * Height - ScaleHeight 870 of a sizable form, f_win); the height of the style differs from one style to another
 * (the vista style of Windows: a pixel or two more, t21 geo), and the client area would not have the size set */
static const int MENU_PX = 20;
class FormMenuBar : public QMenuBar {
public:
    explicit FormMenuBar(QWidget *p) : QMenuBar(p) { setNativeMenuBar(false); }
    QSize sizeHint() const override { return QSize(QMenuBar::sizeHint().width(), MENU_PX); }
    QSize minimumSizeHint() const override { return QSize(QMenuBar::minimumSizeHint().width(), MENU_PX); }
    int heightForWidth(int) const override { return MENU_PX; }
};
class FormWidget : public QWidget {
public:
    vbobj form;
    explicit FormWidget(vbobj f) : QWidget(nullptr), form(f) {}
protected:
    void resizeEvent(QResizeEvent *e) override;
    void closeEvent(QCloseEvent *e) override
    {
        FormState *st = fstate(form);
        if (st->unloading) { e->accept(); return; }
        e->ignore();
        unload_form(form, 0);                   /* vbFormControlMenu */
    }
};

static int border_px(Ctl *c);
static void line_geometry(Ctl *c);
static void align_layout(FormState *st);
void FormWidget::resizeEvent(QResizeEvent *e)
{
    QWidget::resizeEvent(e);
    FormState *st = fstate(form);
    if (st->mdi) align_layout(st);
}
static bool aligning;
static void align_layout(FormState *st);
static double propn(const char *const *props, const char *name, double dflt);
static void ctl_geometry(Ctl *c)
{
    if (c->kind == K_LINE) { line_geometry(c); return; }
    if (c->align && c->form && !aligning) { align_layout(c->form); return; }
    /* a drop-down ComboBox and the DriveListBox: the height comes from the font, not from the .frm */
    if ((c->kind == K_COMBO && (!c->def || propn(c->def->props, "Style", 0) != 1)) || c->kind == K_DRIVE)
        c->height = c->font.size <= 10 ? 315 : std::ceil(QFontMetrics(qfont(c->font)).height() * 1.4) * TPP;
    if (c->w) c->w->setGeometry((int)(c->left / TPP), (int)(c->top / TPP), (int)(c->width / TPP), (int)(c->height / TPP));
    if (c->g) gfx_resize(c->g, (int)(c->width / TPP) - 2 * border_px(c), (int)(c->height / TPP) - 2 * border_px(c));
    if (c->kind == K_USER) uctl_geometry(c);    /* its Resize */
    if (c->ocx) ocx_resized(c);                 /* a CoolBar: its bands */
}

/* ---------------------------------------------------------------- design-time properties */
static bool prop(const char *const *props, const char *name, QString *out)
{
    for (int i = 0; props && props[i]; i += 2)
        if (!strcasecmp(props[i], name)) { if (out) *out = QString::fromUtf8(props[i + 1] + 1); return true; }
    return false;
}
static double propn(const char *const *props, const char *name, double dflt)
{
    QString v;
    if (!prop(props, name, &v)) return dflt;
    bool ok; double d = v.toDouble(&ok);
    return ok ? d : dflt;
}
static long propc(const char *const *props, const char *name, long dflt)
{
    QString v;
    if (!prop(props, name, &v)) return dflt;
    bool ok; double d = v.toDouble(&ok);
    return ok ? norm_color(d) : dflt;
}
static void font_from_props(VFont *f, const char *const *props)
{
    QString v;
    if (prop(props, "Font.Name", &v)) f->name = v;
    f->size = real_size(propn(props, "Font.Size", f->size));
    f->charset = (int)propn(props, "Font.Charset", f->charset);
    f->weight = (int)propn(props, "Font.Weight", f->weight);
    f->bold = f->weight >= 600;
    f->italic = propn(props, "Font.Italic", f->italic ? -1 : 0) != 0;
    f->underline = propn(props, "Font.Underline", f->underline ? -1 : 0) != 0;
    f->strike = propn(props, "Font.Strikethrough", f->strike ? -1 : 0) != 0;
}
static void gfx_from_props(Gfx *g, const char *const *p, double twips_w, double twips_h)
{
    g->pw = (int)(twips_w / TPP); g->ph = (int)(twips_h / TPP);
    g->scalemode = (int)propn(p, "ScaleMode", 1);
    gfx_rescale(g);
    if (g->scalemode == 0) {
        g->sl = propn(p, "ScaleLeft", 0); g->st = propn(p, "ScaleTop", 0);
        g->sw = propn(p, "ScaleWidth", g->pw * TPP); g->sh = propn(p, "ScaleHeight", g->ph * TPP);
    }
    g->autoredraw = propn(p, "AutoRedraw", 0) != 0;
    g->drawwidth = (int)propn(p, "DrawWidth", 1);
    g->drawstyle = (int)propn(p, "DrawStyle", 0);
    g->drawmode = (int)propn(p, "DrawMode", 13);
    g->fillstyle = (int)propn(p, "FillStyle", 1);
    g->fillcolor = propc(p, "FillColor", 0);
    g->fonttransparent = propn(p, "FontTransparent", -1) != 0;
    g->cx = g->cy = 0;
    g->img = QImage();
    gfx_resize(g, g->pw, g->ph);
}

static Kind kind_of(const std::string &cls, std::string *type)
{
    std::string t = cls.substr(cls.find('.') == std::string::npos ? 0 : cls.find('.') + 1);
    *type = t;
    static const std::map<std::string, Kind> m = {
        { "CommandButton", K_COMMAND }, { "Label", K_LABEL }, { "TextBox", K_TEXT }, { "CheckBox", K_CHECK },
        { "OptionButton", K_OPTION }, { "ListBox", K_LIST }, { "ComboBox", K_COMBO }, { "Timer", K_TIMER },
        { "Frame", K_FRAME }, { "PictureBox", K_PICTURE }, { "Menu", K_MENU },
        { "Image", K_IMAGE }, { "Shape", K_SHAPE }, { "Line", K_LINE }, { "HScrollBar", K_HSCROLL }, { "VScrollBar", K_VSCROLL },
        { "DriveListBox", K_DRIVE }, { "DirListBox", K_DIR }, { "FileListBox", K_FILE },
        { "CommonDialog", K_DIALOG }, { "Data", K_DATA } };
    auto it = m.find(t);
    if (it == m.end()) return uctl_class(cls.c_str()) ? K_USER : ocx_type(t) ? K_OCX : K_OTHER;
    return it->second;
}

/* ---------------------------------------------------------------- controls: widgets */
static Ctl *C(vbobj o) { return (Ctl *)vb_oid(o)->ext; }
static QListWidget *LW(Ctl *c) { return qobject_cast<QListWidget *>(c->w.data()); }

/* ListBox Columns: 0 one column scrolled up and down; n > 0 snaking columns (n of them in the width of the list),
 * scrolled sideways (robot: BrionWordSearch sets List1.Columns = 2 ... 6) */
static void list_columns(QListWidget *lw, int n)
{
    if (!lw) return;
    if (n <= 0) {
        lw->setWrapping(false);
        lw->setFlow(QListView::TopToBottom);
        lw->setGridSize(QSize());
        lw->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
        lw->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
        return;
    }
    lw->setFlow(QListView::TopToBottom);
    lw->setWrapping(true);
    lw->setResizeMode(QListView::Adjust);
    lw->setUniformItemSizes(true);
    lw->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    lw->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    int w = lw->viewport()->width();
    if (w < n) w = lw->width() - 2 * lw->frameWidth();
    lw->setGridSize(QSize(qMax(1, w / n), lw->fontMetrics().height() + 2));
}
static QComboBox *CB(Ctl *c) { return qobject_cast<QComboBox *>(c->w.data()); }

static void set_text(Ctl *c, const QString &t, bool fire_change)
{
    if (c->kind == K_TEXT) {
        if (auto *le = qobject_cast<QLineEdit *>(c->w.data())) {
            if (le->text() == t) return;
            QSignalBlocker b(le);
            le->setText(t);
            le->setCursorPosition(0);               /* VB6: SelStart = 0 after Text = ... */
        } else if (auto *pe = qobject_cast<QPlainTextEdit *>(c->w.data())) {
            QString q = QString(t).replace(QStringLiteral("\r\n"), QStringLiteral("\n")).replace('\r', '\n');
            if (pe->toPlainText() == q) return;
            QSignalBlocker b(pe);
            pe->setPlainText(q);
        }
        if (fire_change) fire_ctl(c, "change");
    } else if (c->kind == K_COMBO) {
        auto *cb = qobject_cast<QComboBox *>(c->w.data());
        if (!cb || cb->currentText() == t) return;
        if (!cb->isEditable()) {                /* a Dropdown List: the text must be one of its items */
            int i = cb->findText(t, Qt::MatchFixedString);
            if (i < 0) { if (!t.isEmpty()) vb_raise(383); return; }
            QSignalBlocker b(cb);
            cb->setCurrentIndex(i);
            return;
        }
        QSignalBlocker b(cb);
        QSignalBlocker b2(cb->lineEdit());
        int i = cb->findText(t, Qt::MatchFixedString | Qt::MatchCaseSensitive);
        if (i >= 0) cb->setCurrentIndex(i);    /* the text of an item: that item is the ListIndex, as in VB6 */
        else cb->setCurrentIndex(-1);
        cb->setEditText(t);
        if (fire_change) fire_ctl(c, "change");
    }
}

static QString get_text(Ctl *c)
{
    if (auto *le = qobject_cast<QLineEdit *>(c->w.data())) return le->text();
    if (auto *pe = qobject_cast<QPlainTextEdit *>(c->w.data())) return pe->toPlainText().replace('\n', QStringLiteral("\r\n"));
    if (auto *cb = qobject_cast<QComboBox *>(c->w.data())) return cb->currentText();
    if (auto *lw = qobject_cast<QListWidget *>(c->w.data())) return lw->currentRow() >= 0 ? lw->currentItem()->text() : QString();
    return QString();
}

static void ctl_geometry(Ctl *c);
/* a Label wraps its text at its width; AutoSize without WordWrap makes it one line, as wide as the text */
/* the text a Label shows: with UseMnemonic (the default) & marks the access key, && is an & */
static QString label_shown(Ctl *c)
{
    if (c->def && propn(c->def->props, "UseMnemonic", -1) == 0) return c->caption;
    QString t;
    const QString &s = c->caption;
    for (int i = 0; i < s.size(); i++) {
        if (s[i] == '&') { if (i + 1 < s.size() && s[i + 1] == '&') { t += '&'; i++; } continue; }
        t += s[i];
    }
    return t;
}
static void label_style(Ctl *c, QLabel *l)
{
    if (!l) return;
    l->setAlignment((c->alignment == 1 ? Qt::AlignRight : c->alignment == 2 ? Qt::AlignHCenter : Qt::AlignLeft) | Qt::AlignTop);
    l->setWordWrap(!(c->autosize && !c->wordwrap));
}
/* Label.AutoSize: the size of the caption (WordWrap: the width stays, the height follows the text) */
static void label_fit(Ctl *c)
{
    if (c->kind != K_LABEL || !c->autosize || !c->w) return;
    QFontMetrics fm(qfont(c->font));
    bool wrap = c->wordwrap;
    int b = c->def && propn(c->def->props, "BorderStyle", 0) ? 2 : 0;
    QString t = label_shown(c);
    int wpx, hpx;
    if (!wrap) {
        QStringList lines = t.split('\n');
        wpx = 0;
        for (const QString &ln : lines) wpx = std::max(wpx, fm.horizontalAdvance(QString(ln).remove('\r')));
        hpx = fm.height() * (int)lines.size();
    } else {
        wpx = std::max(1, (int)(c->width / TPP) - b);
        hpx = fm.boundingRect(QRect(0, 0, wpx, 100000), Qt::TextWordWrap, t).height();
    }
    c->width = (wpx + b) * TPP;
    c->height = std::max(hpx + b, fm.height() + b) * TPP;
    ctl_geometry(c);
}

static void apply_look(Ctl *c)
{
    if (!c->w) return;
    if (c->kind == K_SHAPE || c->kind == K_LINE) { c->w->update(); return; }
    if (c->kind == K_USER) return;              /* a user control: its look is its own (ReadProperties) */
    QWidget *w = c->w;
    /* the font again only when it changed: each setFont on a shown widget looks the font up (robot: FontSelector
     * sets ForeColor on 2400 labels, each in a font of its own - a minute) */
    QFont nf = qfont(c->font);
    bool newfont = !c->look_done || !(w->font() == nf) || w->font().pointSizeF() != nf.pointSizeF();
    if (newfont) w->setFont(nf);
    QPalette pal = w->palette();
    QColor b = qcolor(c->back), f = fore_on(c->fore, c->back);
    switch (c->kind) {
    case K_TEXT: case K_LIST: case K_COMBO:
        pal.setColor(QPalette::Base, b); pal.setColor(QPalette::Text, f); break;
    case K_COMMAND:
        /* VB6 paints BackColor only on a graphical button (Style = 1); a standard one keeps the
         * button face of the system (Blanker: BackColor black on a black form) */
        if (c->def && propn(c->def->props, "Style", 0) == 1) pal.setColor(QPalette::Button, b);
        else pal = QApplication::palette();
        break;
    case K_PICTURE:
        break;
    default:
        pal.setColor(QPalette::Window, b); pal.setColor(QPalette::WindowText, f);
        w->setAutoFillBackground(!(c->kind == K_LABEL && c->backstyle == 0));   /* BackStyle 0: transparent */
        break;
    }
    if (!(pal == w->palette())) w->setPalette(pal);
    if (c->look_done && newfont) label_fit(c);  /* a new font; at load the size of the .frm stands */
    c->look_done = true;
}


static Qt::CursorShape pointer_shape(int v)
{
    static const Qt::CursorShape shapes[] = { Qt::ArrowCursor, Qt::ArrowCursor, Qt::CrossCursor, Qt::IBeamCursor, Qt::ArrowCursor,
        Qt::SizeAllCursor, Qt::SizeBDiagCursor, Qt::SizeVerCursor, Qt::SizeFDiagCursor, Qt::SizeHorCursor, Qt::UpArrowCursor,
        Qt::WaitCursor, Qt::ForbiddenCursor, Qt::BusyCursor, Qt::WhatsThisCursor, Qt::SizeAllCursor };
    return v > 0 && v < 16 ? shapes[v] : Qt::ArrowCursor;
}

/* ---------------------------------------------------------------- Picture objects (LoadPicture, .Picture) */
struct Pic { QImage img; int type = 0; };      /* type: 0 none, 1 bitmap, 3 icon */
static std::map<vb_i32, Pic *> pic_handles;    /* the Handle of a Picture: for SelectObject, GetObject (winapi_gui.inc) */
static vb_i32 pic_handle(Pic *p) { vb_i32 h = (vb_i32)(((quintptr)p >> 4) & 0x7FFFFFFF); pic_handles[h] = p; return h; }
static Pic *pic_of(vbobj o);
static const QImage *image_of_value(vbvar v)
{
    if (v.vt != VT_DISPATCH || !v.u.o) return nullptr;
    Pic *p = pic_of(v.u.o);
    return p ? &p->img : nullptr;
}
static vbvar pic_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static void pic_clear(vbobj o)
{
    Pic *p = (Pic *)o->ext;
    for (auto it = pic_handles.begin(); it != pic_handles.end(); ) it = it->second == p ? pic_handles.erase(it) : std::next(it);
    delete p; o->ext = nullptr;
}
static void pic_init(vbobj o) { if (!o->ext) o->ext = new Pic; }
static const vb_class CLS_PIC = { "Picture", (vb_i32)sizeof(vbobjh), pic_init, nullptr, pic_clear, nullptr, 0, "handle",
                                  nullptr, nullptr, nullptr, pic_member, nullptr };
extern "C" const vb_class *vbf_cls_stdpicture(void) { return &CLS_PIC; }
static long himetric(int px) { return std::lround(px * 2540.0 / 96.0); }   /* sizes of pictures: HiMetric */
static vbvar pic_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    Pic *p = (Pic *)vb_oid(o)->ext;
    std::string n = lower(name);
    (void)kind; (void)argc; (void)argv;
    *found = 1;
    if (n == "width") return vb_v_l((vb_i32)(p->type ? himetric(p->img.width()) : 0));
    if (n == "height") return vb_v_l((vb_i32)(p->type ? himetric(p->img.height()) : 0));
    if (n == "type") return vb_v_i((vb_i16)p->type);
    if (n == "handle") return vb_v_l(p->type ? pic_handle(p) : 0);
    if (n == "hpal") return vb_v_l(0);
    *found = 0;
    return vb_v_empty();
}
static vbobj new_pic(const QImage &img, int type)
{
    vbobj o = vb_onew(&CLS_PIC);
    Pic *p = (Pic *)o->ext; p->img = img; p->type = img.isNull() ? 0 : type;
    return o;
}
static Pic *pic_of(vbobj o) { o = vb_oid(o); return o && o->cls == &CLS_PIC ? (Pic *)o->ext : nullptr; }
extern "C" void vbos_ci_path(char *p);
static QString native(const QString &path)
{
#ifdef _WIN32
    return path;
#else
    QString r = path; r.replace('\\', '/');
    QByteArray b = r.toUtf8();
    vbos_ci_path(b.data());                     /* names without case, as on Windows (LoadPicture "Mask.BMP") */
    return QString::fromUtf8(b);
#endif
}
static vbobj loadpicture_hook(vbs path, vb_i32 has)
{
    ensure_app();
    if (!has || path.len == 0) return new_pic(QImage(), 0);
    QString f = native(qs(path));
    if (!QFileInfo::exists(f)) { vb_raise(53); return nullptr; }
    QImage img(f);
    if (img.isNull()) img.load(f, "BMP");                      /* .dib, .rle: bitmaps by another name */
    if (img.isNull()) { vb_raise(481); return nullptr; }       /* Invalid picture */
    QString suf = QFileInfo(f).suffix().toLower();
    return new_pic(img.convertToFormat(img.hasAlphaChannel() ? QImage::Format_ARGB32 : QImage::Format_RGB32), suf == "ico" || suf == "cur" ? 3 : 1);
}
static vbobj picture_data_hook(const char *data, vb_i32 n, vb_i32 type)
{
    ensure_app();
    QImage img = QImage::fromData(QByteArray(data, n), type == 3 ? "ICO" : "BMP");
    if (img.isNull()) return nullptr;
    return new_pic(img.convertToFormat(type == 3 ? QImage::Format_ARGB32 : QImage::Format_RGB32), type);
}
static vbobj picture_get(vbobj *slot);
/* MousePointer: a shape, or 99 = vbCustom: the picture of MouseIcon */
static void apply_pointer(QWidget *w, int mp, vbobj icon)
{
    if (!w) return;
    Pic *p = pic_of(icon);
    if (mp == 99 && p && p->type) w->setCursor(QCursor(QPixmap::fromImage(p->img)));
    else if (mp && mp != 99) w->setCursor(pointer_shape(mp));
    else w->unsetCursor();
}
static vbvar mouseicon_member(vbobj *slot, QWidget *w, int mp, bool set, vbvar val)
{
    if (!set) return vb_vtmp(vb_v_obj(picture_get(slot)));
    vbobj np = val.vt == VT_DISPATCH ? vb_oid(val.u.o) : nullptr;
    if (np) vb_oaddref(np);
    if (*slot) vb_orelease(*slot);
    *slot = np;
    apply_pointer(w, mp, np);
    return vb_v_empty();
}
static void savepicture_hook(vbobj pic, vbs path)
{
    Pic *p = pic_of(pic);
    if (!p || !p->type) { vb_raise(481); return; }
    if (!p->img.save(native(qs(path)), "BMP")) vb_raise(75);
}
static QImage frx_image(const char *const *props, const char *name)
{
    QString v;
    if (!prop(props, name, &v) || v.isEmpty()) return QImage();
    QImage img = QImage::fromData(QByteArray::fromHex(v.toLatin1()));
    /* an icon (or a GIF) keeps its transparent part, as in VB6 */
    return img.convertToFormat(img.hasAlphaChannel() ? QImage::Format_ARGB32 : QImage::Format_RGB32);
}
/* the Picture of a control: 3 (vbPicTypeIcon) for an icon */
static int frx_pictype(const char *const *props, const char *name)
{
    QString v;
    if (!prop(props, name, &v) || v.size() < 8) return 1;
    QString h = v.left(8).toLower();
    return h == "00000100" || h == "00000200" ? 3 : 1;
}

/* ---------------------------------------------------------------- Shape, Line */
class ShapeWidget : public QWidget {
public:
    vbobj self;
    ShapeWidget(vbobj o, QWidget *parent) : QWidget(parent), self(o) { setAttribute(Qt::WA_TransparentForMouseEvents); }
protected:
    void paintEvent(QPaintEvent *) override
    {
        Ctl *c = C(self);
        if (!c) return;
        QPainter p(this);
        int bw = c->borderstyle ? c->borderwidth : 0;
        QRectF r = QRectF(rect()).adjusted(bw / 2.0, bw / 2.0, -bw / 2.0 - (bw ? 0 : 0), -bw / 2.0);
        if (c->shape == 1 || c->shape == 3 || c->shape == 5) {      /* square, circle, rounded square */
            double side = std::min(r.width(), r.height());
            r = QRectF(r.center().x() - side / 2, r.center().y() - side / 2, side, side);
        }
        static const Qt::PenStyle pens[] = { Qt::NoPen, Qt::SolidLine, Qt::DashLine, Qt::DotLine, Qt::DashDotLine, Qt::DashDotDotLine, Qt::SolidLine };
        QPen pen(qcolor(c->bordercolor), std::max(1, c->borderwidth), pens[std::min(6, std::max(0, c->borderstyle))]);
        auto draw = [&](const QBrush &b, const QPen &pn) {
            p.setBrush(b); p.setPen(pn);
            switch (c->shape) {
            case 2: case 3: p.drawEllipse(r); break;
            case 4: case 5: { double rr = std::min(r.width(), r.height()) / 6; p.drawRoundedRect(r, rr, rr); break; }
            default: p.drawRect(r); break;
            }
        };
        if (c->backstyle) draw(QBrush(qcolor(c->back)), Qt::NoPen);
        static const Qt::BrushStyle fills[] = { Qt::SolidPattern, Qt::NoBrush, Qt::HorPattern, Qt::VerPattern, Qt::BDiagPattern, Qt::FDiagPattern, Qt::CrossPattern, Qt::DiagCrossPattern };
        QBrush fill(qcolor(c->fillcolor), fills[std::min(7, std::max(0, c->fillstyle))]);
        draw(fill, c->borderstyle ? pen : QPen(Qt::NoPen));
    }
};

static Gfx *container_gfx(Ctl *c)
{
    if (c->container && C(c->container) && C(c->container)->g) return C(c->container)->g;
    return &c->form->g;
}
class LineWidget : public QWidget {
public:
    vbobj self;
    QPointF a, b;                               /* the ends inside the widget */
    LineWidget(vbobj o, QWidget *parent) : QWidget(parent), self(o) { setAttribute(Qt::WA_TransparentForMouseEvents); }
protected:
    void paintEvent(QPaintEvent *) override
    {
        Ctl *c = C(self);
        if (!c || !c->borderstyle) return;
        QPainter p(this);
        static const Qt::PenStyle pens[] = { Qt::NoPen, Qt::SolidLine, Qt::DashLine, Qt::DotLine, Qt::DashDotLine, Qt::DashDotDotLine, Qt::SolidLine };
        p.setPen(QPen(qcolor(c->bordercolor), std::max(1, c->borderwidth), pens[std::min(6, std::max(0, c->borderstyle))]));
        p.drawLine(a, b);
    }
};
static void line_geometry(Ctl *c)
{
    auto *lw = dynamic_cast<LineWidget *>(c->w.data());
    if (!lw) return;
    Gfx *g = container_gfx(c);
    double ax = px_x(g, c->x1), ay = px_y(g, c->y1), bx = px_x(g, c->x2), by = px_y(g, c->y2);
    int m = std::max(1, c->borderwidth);
    int l = (int)std::floor(std::min(ax, bx)) - m, t = (int)std::floor(std::min(ay, by)) - m;
    int r = (int)std::ceil(std::max(ax, bx)) + m, btm = (int)std::ceil(std::max(ay, by)) + m;
    lw->setGeometry(l, t, r - l + 1, btm - t + 1);
    lw->a = QPointF(ax - l, ay - t); lw->b = QPointF(bx - l, by - t);
    lw->update();
}

/* ---------------------------------------------------------------- scroll bars */
static QScrollBar *SB(Ctl *c) { return qobject_cast<QScrollBar *>(c->w.data()); }
static int sdir(Ctl *c) { return c->smax >= c->smin ? 1 : -1; }
static void scroll_sync(Ctl *c)
{
    QScrollBar *sb = SB(c);
    if (!sb) return;
    QSignalBlocker b(sb);
    sb->setRange(0, std::abs(c->smax - c->smin));
    sb->setSingleStep(c->small); sb->setPageStep(c->large);
    sb->setValue(std::abs(c->sval - c->smin));
}
static void scroll_set(Ctl *c, int v)
{
    if (v < std::min(c->smin, c->smax) || v > std::max(c->smin, c->smax)) { vb_raise(380); return; }
    if (v == c->sval) return;
    c->sval = v;
    scroll_sync(c);
    fire_ctl(c, "change");
}

/* ---------------------------------------------------------------- DriveListBox, DirListBox, FileListBox */
static QString shown_path(const QString &p) { return QDir::toNativeSeparators(p); }
static QStringList sorted_ci(QStringList l) { std::sort(l.begin(), l.end(), [](const QString &a, const QString &b) { int r = a.compare(b, Qt::CaseInsensitive); return r ? r < 0 : a < b; }); return l; }
static QStringList dir_subs(Ctl *c)
{
    QDir d(c->path);
    QStringList subs = d.entryList(QDir::Dirs | QDir::NoDotAndDotDot);
    return sorted_ci(subs);
}
static QStringList dir_parents(Ctl *c)             /* the root first, the current folder last */
{
    QStringList out;
    QDir d(c->path);
    out.prepend(d.absolutePath());
    while (d.cdUp()) out.prepend(d.absolutePath());
    return out;
}
static void dir_fill(Ctl *c)
{
    auto *lw = qobject_cast<QListWidget *>(c->w.data());
    if (!lw) return;
    QSignalBlocker b(lw);
    lw->clear();
    QStringList par = dir_parents(c);
    int depth = 0;
    for (const QString &p : par) {
        QString nm = QDir(p).dirName();
        if (nm.isEmpty()) nm = shown_path(p);
        lw->addItem(QString(depth * 2, ' ') + nm);
        depth++;
    }
    for (const QString &s : dir_subs(c)) lw->addItem(QString(depth * 2, ' ') + s);
    lw->setCurrentRow(par.size() - 1);
}
static void dir_setpath(Ctl *c, const QString &path)
{
    QString p = native(path);
    QFileInfo fi(QDir(c->path).absoluteFilePath(p));
    if (!fi.exists() || !fi.isDir()) { vb_raise(76); return; }   /* Path not found */
    QString np = QDir::cleanPath(fi.absoluteFilePath());
    if (np == c->path) return;
    c->path = np;
    dir_fill(c);
    fire_ctl(c, "change");
}
static QStringList file_names(Ctl *c)
{
    QDir d(c->path);
    QDir::Filters f = QDir::Files | QDir::NoDotAndDotDot;
    if (c->f_hidden) f |= QDir::Hidden;
    if (c->f_system) f |= QDir::System;
    QStringList pats = c->pattern.split(';', Qt::SkipEmptyParts);
    for (auto &x : pats) x = x.trimmed();
    d.setNameFilters(pats.isEmpty() ? QStringList("*") : pats);
    QStringList out;
    for (const QFileInfo &fi : d.entryInfoList(f)) {
        if (!c->f_readonly && !fi.isWritable()) continue;
        out << fi.fileName();
    }
    return sorted_ci(out);
}
static void file_fill(Ctl *c)
{
    auto *lw = qobject_cast<QListWidget *>(c->w.data());
    if (!lw) return;
    QSignalBlocker b(lw);
    lw->clear();
    lw->addItems(file_names(c));
    c->itemdata.assign(lw->count(), 0);
}
static QString drive_of(const QString &root)
{
    QString r = root;
    if (r.size() >= 2 && r[1] == ':') return r.left(2).toLower();   /* Windows: c: */
    return r;
}
static void drive_fill(Ctl *c)
{
    auto *cb = qobject_cast<QComboBox *>(c->w.data());
    if (!cb) return;
    QSignalBlocker b(cb);
    cb->clear();
    QStringList roots;
    for (const QStorageInfo &si : QStorageInfo::mountedVolumes())
        if (si.isValid() && si.isReady()) {
            QString rp = si.rootPath();
            if (rp == "/" || rp.startsWith("/media/") || rp.startsWith("/mnt/") || rp.startsWith("/run/media/") || (rp.size() >= 2 && rp[1] == ':'))
                if (!roots.contains(rp)) roots << rp;
        }
    if (roots.isEmpty()) roots << QDir::rootPath();
    std::sort(roots.begin(), roots.end());
    for (const QString &r : roots) cb->addItem(drive_of(r));
    c->itemdata.assign(cb->count(), 0);
    c->path = drive_of(QDir::rootPath());
    int i = cb->findText(c->path);
    cb->setCurrentIndex(i < 0 ? 0 : i);
}

static const char *const PROP_KEY = "vbctl";
static void make_widget(Ctl *c, QWidget *parent)
{
    QWidget *w = nullptr;
    vbobj self = c->self;
    switch (c->kind) {
    case K_COMMAND: {
        auto *b = new QPushButton(c->caption, parent);
        b->setAutoDefault(false);
        QObject::connect(b, &QPushButton::clicked, [self]() { fire_ctl(C(self), "click"); });
        w = b; break;
    }
    case K_LABEL: {
        auto *l = new QLabel(label_shown(c), parent);
        l->setTextFormat(Qt::PlainText);
        w = l;
        label_style(c, l);
        break;
    }
    case K_TEXT: {
        if (propn(c->def->props, "MultiLine", 0) != 0) {    /* MultiLine: lines, CR LF between them */
            auto *pe = new QPlainTextEdit(parent);
            int sb = (int)propn(c->def->props, "ScrollBars", 0);
            pe->setHorizontalScrollBarPolicy(sb == 1 || sb == 3 ? Qt::ScrollBarAlwaysOn : Qt::ScrollBarAlwaysOff);
            pe->setVerticalScrollBarPolicy(sb == 2 || sb == 3 ? Qt::ScrollBarAlwaysOn : Qt::ScrollBarAlwaysOff);
            pe->setLineWrapMode(sb == 1 || sb == 3 ? QPlainTextEdit::NoWrap : QPlainTextEdit::WidgetWidth);
            pe->setTabChangesFocus(true);
            pe->setReadOnly(propn(c->def->props, "Locked", 0) != 0);
            QObject::connect(pe, &QPlainTextEdit::textChanged, [self]() { fire_ctl(C(self), "change"); });
            w = pe; break;
        }
        auto *le = new QLineEdit(parent);
        QString pc;
        if (prop(c->def->props, "PasswordChar", &pc) && !pc.isEmpty()) le->setEchoMode(QLineEdit::Password);
        int al = (int)propn(c->def->props, "Alignment", 0);
        if (al) le->setAlignment((al == 1 ? Qt::AlignRight : Qt::AlignHCenter) | Qt::AlignVCenter);
        QObject::connect(le, &QLineEdit::textChanged, [self](const QString &) { fire_ctl(C(self), "change"); });
        w = le; break;
    }
    case K_CHECK: {
        auto *cb = new QCheckBox(c->caption, parent);
        QObject::connect(cb, &QCheckBox::clicked, [self](bool on) { Ctl *x = C(self); x->value = on ? 1 : 0; fire_ctl(x, "click"); });
        w = cb; break;
    }
    case K_OPTION: {
        auto *rb = new QRadioButton(c->caption, parent);
        QObject::connect(rb, &QRadioButton::clicked, [self](bool) { fire_ctl(C(self), "click"); });
        w = rb; break;
    }
    case K_LIST: {
        auto *lw = new QListWidget(parent);
        int ms = c->def ? (int)propn(c->def->props, "MultiSelect", 0) : 0;   /* 1 Simple, 2 Extended */
        if (ms) lw->setSelectionMode(ms == 1 ? QAbstractItemView::MultiSelection : QAbstractItemView::ExtendedSelection);
        QObject::connect(lw, &QListWidget::itemClicked, [self](QListWidgetItem *) { fire_ctl(C(self), "click"); });
        int cols = c->def ? (int)propn(c->def->props, "Columns", 0) : 0;
        if (cols > 0) {                          /* after the list has its size and font */
            QPointer<QListWidget> pl(lw);
            QTimer::singleShot(0, lw, [pl, cols]() { if (pl) list_columns(pl, cols); });
        }
        w = lw; break;
    }
    case K_COMBO: {
        auto *cb = new QComboBox(parent);
        bool list_only = c->def && propn(c->def->props, "Style", 0) == 2;   /* 2 - Dropdown List: no text box */
        cb->setEditable(!list_only);
        cb->setInsertPolicy(QComboBox::NoInsert);
        QObject::connect(cb, QOverload<int>::of(&QComboBox::activated), [self](int) { fire_ctl(C(self), "click"); });
        if (cb->lineEdit())
            QObject::connect(cb->lineEdit(), &QLineEdit::textChanged, [self](const QString &) { fire_ctl(C(self), "change"); });
        w = cb; break;
    }
    case K_FRAME: w = new QGroupBox(c->caption, parent); break;
    case K_PICTURE: {
        c->g = new Gfx;
        c->g->font = &c->font;
        c->g->back = c->back; c->g->fore = c->fore;
        bool border = propn(c->def->props, "BorderStyle", 1) != 0;
        double bw = border ? 4 * TPP : 0;
        gfx_from_props(c->g, c->def->props, c->width - bw, c->height - bw);
        auto *cv = new Canvas(c->g, parent);
        if (border) { cv->setFrameStyle(QFrame::Panel | QFrame::Sunken); cv->setLineWidth(2); }
        cv->setFocusPolicy(Qt::StrongFocus);
        w = cv; break;
    }
    case K_TIMER: {
        c->timer = new QTimer();
        QObject::connect(c->timer, &QTimer::timeout, [self]() { Ctl *x = C(self); if (x->form && x->enabled) fire_ctl(x, "timer"); });
        return;
    }
    case K_MENU: return;                       /* made by build_menus */
    case K_USER:
        w = uctl_make(c, parent);
        break;
    case K_DATA:
        data_load(c, c->def ? c->def->props : nullptr);
        w = data_make(c, parent);
        break;
    case K_OCX:
        c->ocx = ocx_new(c);
        w = ocx_make(c, parent);
        if (!w) return;                         /* ImageList, PictureClip: no window */
        break;
    case K_DIALOG: {                            /* no window: the dialogs come with Show... */
        auto *d = c->dlg = new CDlg;
        const char *const *p = c->def->props;
        QString v;
        if (prop(p, "DialogTitle", &v)) d->title = v;
        if (prop(p, "Filter", &v)) d->filter = v;
        if (prop(p, "FileName", &v)) d->filename = v;
        if (prop(p, "InitDir", &v)) d->initdir = v;
        if (prop(p, "DefaultExt", &v)) d->defaultext = v;
        if (prop(p, "FontName", &v)) d->fontname = v;
        d->filterindex = (int)propn(p, "FilterIndex", 0);
        d->flags = (long)propn(p, "Flags", 0);
        d->cancelerror = propn(p, "CancelError", 0) != 0;
        d->color = (long)propn(p, "Color", 0);
        d->fontsize = propn(p, "FontSize", 8.25);
        d->min = (int)propn(p, "Min", 0); d->max = (int)propn(p, "Max", 0);
        if (prop(p, "HelpFile", &v)) d->helpfile = v;
        if (prop(p, "HelpKey", &v)) d->helpkey = v;
        d->helpcommand = (vb_i32)propn(p, "HelpCommand", 0); d->helpcontext = (vb_i32)propn(p, "HelpContext", 0);
        return;
    }
    case K_IMAGE: { auto *l = new QLabel(parent); l->setScaledContents(c->stretch); w = l; break; }
    case K_SHAPE: w = new ShapeWidget(self, parent); break;
    case K_LINE: w = new LineWidget(self, parent); break;
    case K_HSCROLL: case K_VSCROLL: {
        auto *sb = new QScrollBar(c->kind == K_HSCROLL ? Qt::Horizontal : Qt::Vertical, parent);
        sb->setFocusPolicy(Qt::StrongFocus);
        QObject::connect(sb, &QScrollBar::valueChanged, [self](int v) {
            Ctl *x = C(self);
            int nv = x->smin + sdir(x) * v;
            if (nv == x->sval) return;
            x->sval = nv;
            QScrollBar *s = SB(x);
            fire_ctl(x, s && s->isSliderDown() ? "scroll" : "change");
        });
        QObject::connect(sb, &QScrollBar::sliderReleased, [self]() { fire_ctl(C(self), "change"); });
        w = sb; break;
    }
    case K_DRIVE: {
        auto *cb = new QComboBox(parent);
        QObject::connect(cb, QOverload<int>::of(&QComboBox::activated), [self](int i) {
            Ctl *x = C(self); auto *b = CB(x);
            if (!b || i < 0) return;
            if (b->itemText(i) != x->path) { x->path = b->itemText(i); fire_ctl(x, "change"); }
        });
        w = cb; break;
    }
    case K_DIR: {
        auto *lw = new QListWidget(parent);
        QObject::connect(lw, &QListWidget::itemClicked, [self](QListWidgetItem *) { fire_ctl(C(self), "click"); });
        QObject::connect(lw, &QListWidget::itemDoubleClicked, [self](QListWidgetItem *it) {
            Ctl *x = C(self); auto *l = qobject_cast<QListWidget *>(x->w.data());
            if (!l) return;
            int row = l->row(it);
            QStringList par = dir_parents(x), subs = dir_subs(x);
            QString to = row < par.size() ? par[row] : QDir(x->path).absoluteFilePath(subs.value(row - par.size()));
            dir_setpath(x, to);
        });
        w = lw; break;
    }
    case K_FILE: {
        auto *lw = new QListWidget(parent);
        QObject::connect(lw, &QListWidget::itemClicked, [self](QListWidgetItem *) { fire_ctl(C(self), "click"); });
        w = lw; break;
    }
    default: w = new QWidget(parent); break;
    }
    c->w = w;
    if (!c->tooltip.isEmpty()) w->setToolTip(c->tooltip);
    if (c->kind == K_HSCROLL || c->kind == K_VSCROLL) scroll_sync(c);
    if (c->def) {                               /* MousePointer and MouseIcon of the designer */
        QImage mi = frx_image(c->def->props, "MouseIcon");
        if (!mi.isNull() && !c->mouseicon) { c->mouseicon = new_pic(mi, 3); vb_oaddref(c->mouseicon); }
    }
    if (c->mousepointer) apply_pointer(w, c->mousepointer, c->mouseicon);
    w->setProperty(PROP_KEY, QVariant::fromValue((qulonglong)(quintptr)self));
    ctl_geometry(c);
    apply_look(c);
    w->setEnabled(c->enabled);
    w->setVisible(c->visible);
}

static void timer_update(Ctl *c)
{
    if (!c->timer) return;
    if (c->enabled && c->interval > 0 && c->form && c->form->loaded) c->timer->start(c->interval);
    else c->timer->stop();
}

/* ---------------------------------------------------------------- the Font object */
struct FontRef { vbobj owner; VFont own; };     /* owner 0: a font of its own (New StdFont) */
static VFont *font_of(vbobj owner);
static void font_changed(vbobj owner);
static vbvar font_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    FontRef *r = (FontRef *)vb_oid(o)->ext;
    VFont *f = r->owner ? font_of(r->owner) : &r->own;
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!f) { vb_raise(91); return vb_v_empty(); }
    if (n == "name") { if (!set) return vstr(f->name); f->name = qs(vb_v2s(val)); }
    else if (n == "size") { if (!set) return vb_v_c((vb_cy)std::llround(f->size * 10000)); double s = vb_v2d(val); if (s <= 0 || s > 2160) { vb_raise(380); return vb_v_empty(); } f->size = real_size(s); }
    else if (n == "bold") { if (!set) return vb_v_bool(f->bold ? -1 : 0); f->bold = vb_v2bool(val) != 0; f->weight = f->bold ? 700 : 400; }
    else if (n == "italic") { if (!set) return vb_v_bool(f->italic ? -1 : 0); f->italic = vb_v2bool(val) != 0; }
    else if (n == "underline") { if (!set) return vb_v_bool(f->underline ? -1 : 0); f->underline = vb_v2bool(val) != 0; }
    else if (n == "strikethrough") { if (!set) return vb_v_bool(f->strike ? -1 : 0); f->strike = vb_v2bool(val) != 0; }
    else if (n == "weight") { if (!set) return vb_v_i((vb_i16)f->weight); f->weight = vb_v2i(val) > 550 ? 700 : 400; f->bold = f->weight == 700; }
    else if (n == "charset") { if (!set) return vb_v_i((vb_i16)f->charset); f->charset = vb_v2i(val); }
    else if (n == "hfont" && !set) return vb_v_l(font_handle(*f));    /* IFont.hFont: for SelectObject */
    else { *found = 0; return vb_v_empty(); }
    if (r->owner) font_changed(r->owner);
    return vb_v_empty();
}
static void font_clear(vbobj o) { FontRef *r = (FontRef *)o->ext; if (r) { if (r->owner) vb_orelease(r->owner); delete r; } o->ext = nullptr; }
static void font_init(vbobj o) { if (!o->ext) o->ext = new FontRef{ nullptr, VFont() }; }
static const vb_class CLS_FONT = { "StdFont", (vb_i32)sizeof(vbobjh), font_init, nullptr, font_clear, nullptr, 0, "name",
                                   nullptr, nullptr, nullptr, font_member, nullptr };
static vbvar font_object(vbobj owner)
{
    vbobj o = vb_onew(&CLS_FONT);
    FontRef *r = (FontRef *)o->ext;
    r->owner = owner; vb_oaddref(owner);
    return vb_vtmp(vb_v_obj(o));
}
/* New StdFont, New StdPicture (the compiler: Dim f As New StdFont) */
extern "C" const vb_class *vbf_cls_stdfont(void) { return &CLS_FONT; }
/* FontName, FontSize, FontBold ... of forms and controls */
static vbvar font_prop(vbobj owner, VFont *f, const std::string &n, bool set, vbvar val, int *found)
{
    *found = 1;
    if (n == "font") {
        if (!set) return font_object(owner);
        if (val.vt != VT_DISPATCH) {            /* Text1.Font = "Arial": the font's default property (Name) */
            f->name = qs(vb_v2s(val));
            font_changed(owner);
            return vb_v_empty();
        }
        /* Set Text1.Font = f: the control takes the font's settings */
        vbobj fo = val.vt == VT_DISPATCH ? vb_oid(val.u.o) : nullptr;
        if (!fo || fo->cls != &CLS_FONT) { vb_raise(13); return vb_v_empty(); }
        FontRef *r = (FontRef *)fo->ext;
        VFont *src = r->owner ? font_of(r->owner) : &r->own;
        if (src && src != f) *f = *src;
        font_changed(owner);
        return vb_v_empty();
    }
    if (n == "fontname") { if (!set) return vstr(f->name); f->name = qs(vb_v2s(val)); }
    else if (n == "fontsize") { if (!set) return vb_v_f((float)f->size); double s = vb_v2d(val); if (s <= 0 || s > 2160) { vb_raise(380); return vb_v_empty(); } f->size = real_size(s); }
    else if (n == "fontbold") { if (!set) return vb_v_bool(f->bold ? -1 : 0); f->bold = vb_v2bool(val) != 0; f->weight = f->bold ? 700 : 400; }
    else if (n == "fontitalic") { if (!set) return vb_v_bool(f->italic ? -1 : 0); f->italic = vb_v2bool(val) != 0; }
    else if (n == "fontunderline") { if (!set) return vb_v_bool(f->underline ? -1 : 0); f->underline = vb_v2bool(val) != 0; }
    else if (n == "fontstrikethru") { if (!set) return vb_v_bool(f->strike ? -1 : 0); f->strike = vb_v2bool(val) != 0; }
    else { *found = 0; return vb_v_empty(); }
    font_changed(owner);
    return vb_v_empty();
}

/* ---------------------------------------------------------------- focus */
static Ctl *ctl_of_widget(QObject *w)
{
    for (; w; w = w->parent()) {
        QVariant v = w->property(PROP_KEY);
        if (v.isValid()) { vbobj o = (vbobj)(quintptr)v.toULongLong(); return C(o); }
        if (dynamic_cast<FormWidget *>(w)) return nullptr;
    }
    return nullptr;
}
static FormState *form_of_widget(QObject *o)
{
    for (auto *w = qobject_cast<QWidget *>(o); w; w = w->parentWidget()) {
        if (auto *fw = dynamic_cast<FormWidget *>(w)) return fstate(fw->form);
        if (w->isWindow()) return nullptr;      /* a dialog (MsgBox ...) or another window */
    }
    return nullptr;
}
static QWidget *focus_widget(Ctl *c)
{
    if (!c || !c->w) return nullptr;
    if (c->kind == K_USER) return uctl_focus_widget(c);
    if (auto *cb = qobject_cast<QComboBox *>(c->w.data())) return cb->lineEdit() ? (QWidget *)cb->lineEdit() : (QWidget *)cb;
    return c->w;
}
static bool can_focus(Ctl *c)
{
    if (!c->w || !c->visible || !c->enabled) return false;
    if (c->kind == K_USER) {                    /* its own controls, or itself (CanGetFocus) */
        if (!uctl_can_focus(c)) return false;
        for (vbobj p = c->container; p; p = C(p)->container) if (!C(p)->visible || !C(p)->enabled) return false;
        return true;
    }
    switch (c->kind) {
    case K_COMMAND: case K_TEXT: case K_CHECK: case K_OPTION: case K_LIST: case K_COMBO: case K_PICTURE:
    case K_HSCROLL: case K_VSCROLL: case K_DRIVE: case K_DIR: case K_FILE: break;
    case K_OCX: if (!ocx_focusable(c->ocx)) return false; break;
    default: return false;
    }
    if (propn(c->def->props, "TabStop", -1) == 0 && c->kind != K_PICTURE) return true;
    for (vbobj p = c->container; p; p = C(p)->container) if (!C(p)->visible || !C(p)->enabled) return false;
    return true;
}
static Ctl *first_focusable(FormState *st)
{
    Ctl *best = nullptr;
    for (vbobj o : st->ctls) { Ctl *c = C(o); if (can_focus(c) && (!best || c->tabindex < best->tabindex)) best = c; }
    return best;
}
/* VB6 changes ActiveControl at once; LostFocus and GotFocus come from the message loop (f_ui) */
/* the Validate event of the control that has the focus; true: it may lose it */
static bool validate(Ctl *c)
{
    if (!c || !c->form) return true;
    vbvar a[1] = { vb_v_bool(0) };
    fire_ctl(c, "validate", 1, a);
    return !vb_v2bool(a[0]);
}
static bool focus_to(FormState *st, Ctl *c, bool events = true, bool user = false)
{
    vbobj old = st->focus, now = c ? c->self : nullptr;
    if (old == now) return true;
    /* Validate: when the user moves the focus (Tab, a click), not with SetFocus (f_misc) */
    if (events && user && old && c && c->causesval && !validate(C(old))) {   /* Cancel: the focus stays */
        if (QWidget *w = focus_widget(C(old))) w->setFocus(Qt::OtherFocusReason);
        return false;
    }
    /* a user control: the control in it that has the focus (or itself), and in it the events
     * VB6 sends (oracle f_uctl2): the old one's LostFocus and ExitFocus, the new one's EnterFocus,
     * then the form's LostFocus / GotFocus of the two, then the new one's own GotFocus */
    Ctl *oc = old ? C(old) : nullptr;
    vbobj old_uc = oc && oc->kind == K_USER ? oc->uc : nullptr, new_uc = c && c->kind == K_USER ? c->uc : nullptr;
    Ctl *oin = old_uc ? uctl_inner_focus(oc) : nullptr;
    Ctl *nin = new_uc ? uctl_inner_focus(c) : nullptr;
    vbobj old_in = oin ? oin->self : nullptr, new_in = nin ? nin->self : nullptr;
    if (new_uc) fstate(new_uc)->focus = new_in;
    st->focus = now;
    if (QWidget *w = focus_widget(c)) w->setFocus(Qt::OtherFocusReason);
    if (!events) return true;
    for (vbobj x : { old, now, old_uc, new_uc, old_in, new_in }) if (x) vb_oaddref(x);
    post([old, now, old_uc, new_uc, old_in, new_in]() {
        if (old_uc) {
            if (old_in) { Ctl *i = C(old_in); if (i && i->form) fire_ctl(i, "lostfocus"); }
            else fire_form(old_uc, "lostfocus");
            fire_form(old_uc, "exitfocus");
        }
        if (new_uc) fire_form(new_uc, "enterfocus");
        if (old) { Ctl *o = C(old); if (o && o->form) fire_ctl(o, "lostfocus"); }
        if (now) { Ctl *n = C(now); if (n && n->form) fire_ctl(n, "gotfocus"); }
        if (new_uc) {
            if (new_in) { Ctl *i = C(new_in); if (i && i->form) fire_ctl(i, "gotfocus"); }
            else fire_form(new_uc, "gotfocus");
        }
        for (vbobj x : { old, now, old_uc, new_uc, old_in, new_in }) if (x) vb_orelease(x);
    });
    return true;
}
/* Tab / Shift+Tab: the next control in the tab order (the keys themselves have no events) */
static void tab_move(FormState *st, bool back)
{
    std::vector<Ctl *> v;
    for (vbobj o : st->ctls) { Ctl *c = C(o); if (can_focus(c) && propn(c->def->props, "TabStop", -1) != 0) v.push_back(c); }
    if (v.empty()) return;
    std::stable_sort(v.begin(), v.end(), [](Ctl *a, Ctl *b) { return a->tabindex < b->tabindex; });
    int cur = -1;
    for (size_t i = 0; i < v.size(); i++) if (st->focus == v[i]->self) cur = (int)i;
    int n = (int)v.size();
    int next = cur < 0 ? 0 : back ? (cur + n - 1) % n : (cur + 1) % n;
    focus_to(st, v[next], true, true);
}

/* ---------------------------------------------------------------- keyboard */
static int bypass;                               /* > 0: key events that go to Qt as they are */
static int vk_of_qt(int key, const QString &text)
{
    if (key >= Qt::Key_A && key <= Qt::Key_Z) return 'A' + (key - Qt::Key_A);
    if (key >= Qt::Key_0 && key <= Qt::Key_9) return '0' + (key - Qt::Key_0);
    if (key >= Qt::Key_F1 && key <= Qt::Key_F24) return 112 + (key - Qt::Key_F1);
    switch (key) {
    case Qt::Key_Backspace: return 8; case Qt::Key_Tab: case Qt::Key_Backtab: return 9; case Qt::Key_Return: case Qt::Key_Enter: return 13;
    case Qt::Key_Shift: return 16; case Qt::Key_Control: return 17; case Qt::Key_Alt: return 18; case Qt::Key_Pause: return 19;
    case Qt::Key_CapsLock: return 20; case Qt::Key_Escape: return 27; case Qt::Key_Space: return 32; case Qt::Key_PageUp: return 33;
    case Qt::Key_PageDown: return 34; case Qt::Key_End: return 35; case Qt::Key_Home: return 36; case Qt::Key_Left: return 37;
    case Qt::Key_Up: return 38; case Qt::Key_Right: return 39; case Qt::Key_Down: return 40; case Qt::Key_Print: return 44;
    case Qt::Key_Insert: return 45; case Qt::Key_Delete: return 46; case Qt::Key_Help: return 47; case Qt::Key_NumLock: return 144;
    case Qt::Key_ScrollLock: return 145;
    default: break;
    }
    if (text.size() == 1) {                     /* the US keyboard */
        static const char *plain = ";=,-./`[\\]'", *shifted = ":+<_>?~{|}\"";
        static const int codes[] = { 186, 187, 188, 189, 190, 191, 192, 219, 220, 221, 222 };
        static const char *digits_sh = ")!@#$%^&*(";
        char ch = text[0].toLatin1();
        if (ch) {
            if (const char *p = strchr(plain, ch)) return codes[p - plain];
            if (const char *p = strchr(shifted, ch)) return codes[p - shifted];
            if (const char *p = strchr(digits_sh, ch)) return '0' + (int)(p - digits_sh);
        }
    }
    return 0;
}
static int shift_of(Qt::KeyboardModifiers m)
{
    return ((m & Qt::ShiftModifier) ? 1 : 0) | ((m & Qt::ControlModifier) ? 2 : 0) | ((m & Qt::AltModifier) ? 4 : 0);
}
static int ascii_of(const QString &text, Qt::KeyboardModifiers m, int vk)
{
    if ((m & Qt::ControlModifier) && vk >= 'A' && vk <= 'Z') return vk & 31;
    if (text.isEmpty()) return 0;
    ushort u = text[0].unicode();
    if (u < 32 || u == 127) return u;
    vb_i32 mk = vb_mark();
    int a = vb_asc(vs(QString(text[0])));
    vb_release(mk);
    return a & 0xFF;
}
static QString char_of(int ascii)
{
    if (ascii < 32) return QString(QChar((ushort)ascii));
    vb_i32 mk = vb_mark();
    QString r = qs(vb_chr(ascii));
    vb_release(mk);
    return r;
}
/* KeyDown / KeyUp: the form first when KeyPreview is on; KeyCode = 0 stops it */
static int key_events(FormState *st, Ctl *c, const char *ev, int code, int shift)
{
    if (st->keypreview || !c) {
        vbvar a[2] = { vb_v_i((vb_i16)code), vb_v_i((vb_i16)shift) };
        fire_form(st->form, ev, 2, a);
        code = vb_v2i(a[0]);
    }
    if (c && code && c->form) {
        vbvar a[2] = { vb_v_i((vb_i16)code), vb_v_i((vb_i16)shift) };
        fire_ctl(c, ev, 2, a);
        code = vb_v2i(a[0]);
    }
    return code;
}
static int keypress_events(FormState *st, Ctl *c, int ascii)
{
    if (st->keypreview || !c) {
        vbvar a[1] = { vb_v_i((vb_i16)ascii) };
        fire_form(st->form, "keypress", 1, a);
        ascii = vb_v2i(a[0]);
    }
    if (c && ascii && c->form) {
        vbvar a[1] = { vb_v_i((vb_i16)ascii) };
        fire_ctl(c, "keypress", 1, a);
        ascii = vb_v2i(a[0]);
    }
    return ascii;
}
static void deliver(QWidget *w, QEvent::Type t, int key, Qt::KeyboardModifiers m, const QString &text, bool rep = false)
{
    if (!w) return;
    QPointer<QWidget> p(w);
    QKeyEvent e(t, key, m, text, rep);
    bypass++;
    QCoreApplication::sendEvent(w, &e);
    bypass--;
}
/* the keys of a scroll bar (Windows): arrows SmallChange, PgUp/PgDn LargeChange, Home/End Min/Max */
static bool scroll_key(Ctl *c, int vk)
{
    int pos = std::abs(c->sval - c->smin), range = std::abs(c->smax - c->smin);
    switch (vk) {
    case 37: case 38: pos -= c->small; break;
    case 39: case 40: pos += c->small; break;
    case 33: pos -= c->large; break;
    case 34: pos += c->large; break;
    case 36: pos = 0; break;
    case 35: pos = range; break;
    default: return false;
    }
    pos = std::max(0, std::min(range, pos));
    int v = c->smin + sdir(c) * pos;
    if (v != c->sval) { c->sval = v; scroll_sync(c); fire_ctl(c, "change"); }
    return true;
}

/* help.inc */
static void help_f1(FormState *st, Ctl *c);
static bool help_menu_f1(QMenu *m);
static bool help_whatsthis_event(QObject *obj, QEvent *e);
static void help_whatsthis_mode();
static void whatsthis_ctl(Ctl *c, QPoint pt);
static vb_i32 ctl_helpid(Ctl *c, const char *name);
static void dlg_showhelp(Ctl *c);
static void msgbox_help();

/* one key: the VB6 events, then what the control does with what is left of the key */
static void key_down(FormState *st, int key, Qt::KeyboardModifiers m, const QString &text, int vk, bool rep = false)
{
    if (vk == 9 && !(m & (Qt::ControlModifier | Qt::AltModifier))) { tab_move(st, (m & Qt::ShiftModifier) || key == Qt::Key_Backtab); return; }
    if ((m & Qt::AltModifier) && !(m & Qt::ControlModifier) && text.size() == 1 && text[0].isLetterOrNumber()
        && uctl_access_key(st, text[0].toLower().unicode())) return;
    if ((m & Qt::AltModifier) && !(m & Qt::ControlModifier) && key >= Qt::Key_A && key <= Qt::Key_Z
        && uctl_access_key(st, 'a' + (key - Qt::Key_A))) return;
    Ctl *c = st->focus ? C(st->focus) : nullptr;
    QWidget *w = c ? focus_widget(c) : (QWidget *)st->canvas.data();
    QPointer<QWidget> pw(w);
    int code = key_events(st, c, "keydown", vk, shift_of(m));
    if (vb_errpend) return;
    QString t = text;
    int ascii = ascii_of(text, m, vk);
    bool pass = true;
    if (ascii) {
        int na = keypress_events(st, c, ascii);
        if (vb_errpend) return;
        if (!na) pass = false;
        else if (na != ascii) t = char_of(na);
    } else if (!code) pass = false;
    if (pass && vk == 112 && st->loaded && !(m & (Qt::ControlModifier | Qt::AltModifier))) {   /* F1: the help (help.inc) */
        if (!(m & Qt::ShiftModifier)) help_f1(st, c);
        else if (st->whatsthis) help_whatsthis_mode();             /* Shift+F1: What's This */
        return;
    }
    if (pass && pw && st->loaded) {
        if (c && (c->kind == K_HSCROLL || c->kind == K_VSCROLL) && scroll_key(c, vk)) return;
        deliver(pw, QEvent::KeyPress, key, m, t, rep);
    }
}
static void key_up(FormState *st, int key, Qt::KeyboardModifiers m, const QString &text, int vk)
{
    if (vk == 9 && !(m & (Qt::ControlModifier | Qt::AltModifier))) return;
    Ctl *c = st->focus ? C(st->focus) : nullptr;
    QWidget *w = c ? focus_widget(c) : (QWidget *)st->canvas.data();
    QPointer<QWidget> pw(w);
    key_events(st, c, "keyup", vk, shift_of(m));
    if (!vb_errpend && pw && st->loaded) deliver(pw, QEvent::KeyRelease, key, m, text);
}

/* ---------------------------------------------------------------- mouse
 * VB6: MouseDown, MouseUp, Click; a double click: MouseDown, MouseUp, Click, DblClick, MouseUp.
 * X and Y are in the scale of the form / PictureBox, for other controls in the scale of
 * their container. Click comes from here for the controls that have no Qt click signal. */
static int vb_button(Qt::MouseButtons b)
{
    return ((b & Qt::LeftButton) ? 1 : 0) | ((b & Qt::RightButton) ? 2 : 0) | ((b & Qt::MiddleButton) ? 4 : 0);
}
static int mouse_bypass;
static QPoint cursor_pos;                        /* SetCursorPos / the last mouse event */
static QWidget *widget_at(QPoint gp);

/* X, Y of a mouse or drag event: the scale of a form or PictureBox, for other controls the
 * scale of their container, from the top left corner of the control */
static void coords_of(FormState *st, Ctl *c, QPointF gp, double *x, double *y)
{
    QWidget *cw = c ? c->w.data() : (QWidget *)st->canvas.data();
    QPointF lp = cw ? map_from_global(cw, gp) : QPointF();
    if (c && c->g) { int b = border_px(c); *x = c->g->sl + un_w(c->g, lp.x() - b); *y = c->g->st + un_h(c->g, lp.y() - b); }
    else if (!c) { *x = st->g.sl + un_w(&st->g, lp.x()); *y = st->g.st + un_h(&st->g, lp.y()); }
    else { Gfx *g = container_gfx(c); *x = un_w(g, lp.x()); *y = un_h(g, lp.y()); }
}

/* ---------------------------------------------------------------- drag and drop (DragMode, Drag)
 * VB6's own drag: while a control is dragged, DragOver goes to what is under the cursor
 * (State 0 enter, 1 leave, 2 over) and DragDrop to where it is dropped; no mouse events. */
static vbobj drag_src;                           /* the control being dragged */
static vbobj drag_tgt;                           /* the control or form under it (weak) */
static void target_at(QPoint gp, FormState **tst, Ctl **tc)
{
    *tst = nullptr; *tc = nullptr;
    QWidget *w = widget_at(gp);
    if (!w) return;
    FormState *st = form_of_widget(w);
    if (!st || !st->loaded) return;
    Ctl *c = ctl_of_widget(w);
    if (c && !c->form) c = nullptr;
    *tst = st; *tc = c;
}
static void drag_event(FormState *st, Ctl *c, const char *ev, QPoint gp, int state)
{
    double x, y;
    coords_of(st, c, QPointF(gp), &x, &y);
    vbvar a[4] = { vb_vtmp(vb_v_obj(drag_src)), vb_v_f((float)x), vb_v_f((float)y), vb_v_i((vb_i16)state) };
    int n = state < 0 ? 3 : 4;
    if (c) fire_ctl(c, ev, n, a); else fire_form(st->form, ev, n, a);
}
static void drag_move(QPoint gp)
{
    FormState *tst; Ctl *tc;
    target_at(gp, &tst, &tc);
    vbobj t = tc ? tc->self : tst ? tst->form : nullptr;
    if (t != drag_tgt) {
        if (drag_tgt) {                          /* leave the old target */
            vbobj o = vb_oid(drag_tgt);
            if (o->cls->form) { if (fstate(o)->loaded) drag_event(fstate(o), nullptr, "dragover", gp, 1); }
            else if (C(o) && C(o)->form) drag_event(C(o)->form, C(o), "dragover", gp, 1);
        }
        drag_tgt = t;
        if (t) drag_event(tst, tc, "dragover", gp, 0);
    } else if (t) drag_event(tst, tc, "dragover", gp, 2);
}
static void drag_end(QPoint gp, bool drop)
{
    if (!drag_src) return;
    FormState *tst; Ctl *tc;
    target_at(gp, &tst, &tc);
    if (drop) { if (tst) drag_event(tst, tc, "dragdrop", gp, -1); }
    else if (drag_tgt) drag_move(QPoint(-100000, -100000));   /* cancel: leave */
    vb_orelease(drag_src);
    drag_src = nullptr; drag_tgt = nullptr;
}
static void drag_begin(Ctl *c)
{
    if (drag_src) drag_end(cursor_pos, false);
    drag_src = c->self; vb_oaddref(drag_src);
    /* the first DragOver goes to the dragged control itself, wherever the cursor is (f_misc) */
    drag_tgt = c->self;
    drag_event(c->form, c, "dragover", cursor_pos, 0);
}

static bool mouse_event(QObject *obj, QMouseEvent *me)
{
    if (mouse_bypass) return false;
    /* Qt passes an event that a widget ignores on to its parent as a new event object:
     * the same type, time and place is the same event (VB6: only the control gets it) */
    static QEvent::Type last_t; static quint64 last_ts; static QPointF last_gp;
    if (me->type() == last_t && me->timestamp() == last_ts && ev_gpos(me) == last_gp) return false;
    auto *w = qobject_cast<QWidget *>(obj);
    if (!w) return false;
    FormState *st = form_of_widget(w);
    if (!st || !st->loaded || !st->canvas || !(w == st->canvas || st->canvas->isAncestorOf(w))) return false;
    last_t = me->type(); last_ts = me->timestamp(); last_gp = ev_gpos(me);
    cursor_pos = ev_gpos(me).toPoint();
    Ctl *c = ctl_of_widget(w);
    if (c && !c->form) return false;
    if (c && c->kind == K_OCX && ocx_eat_mouse(c->ocx, me)) return true;   /* a CoolBar: a double click on a gripper */
    if (drag_src) {                             /* dragging: DragOver / DragDrop instead of mouse events */
        if (me->type() == QEvent::MouseMove) drag_move(cursor_pos);
        else if (me->type() == QEvent::MouseButtonRelease) { drag_move(cursor_pos); drag_end(cursor_pos, true); }   /* f_misc: over, then drop */
        return true;
    }
    if (c && c->dragmode == 1 && me->type() == QEvent::MouseButtonPress && me->button() == Qt::LeftButton) {
        drag_begin(c);                          /* DragMode = vbAutomatic: no MouseDown */
        return true;
    }
    QWidget *cw = c ? c->w.data() : (QWidget *)st->canvas.data();
    if (!cw) return false;
    QPointF lp = map_from_global(cw, ev_gpos(me));
    double x, y;
    coords_of(st, c, ev_gpos(me), &x, &y);
    int button = me->type() == QEvent::MouseMove ? vb_button(me->buttons()) : vb_button(me->button());
    vbvar a[4] = { vb_v_i((vb_i16)button), vb_v_i((vb_i16)shift_of(me->modifiers())), vb_v_f((float)x), vb_v_f((float)y) };
    bool own_click = !c || c->kind == K_LABEL || c->kind == K_IMAGE || c->kind == K_PICTURE || c->kind == K_FRAME ||
                     (c->kind == K_OCX && ocx_clicks(c->ocx));
    bool &dbl = c ? c->dbl : st->dbl;
    auto ev = [&](const char *name, int n) { if (c) fire_ctl(c, name, n, a); else fire_form(st->form, name, n, a); };
    switch (me->type()) {
    case QEvent::MouseButtonPress: dbl = false; ev("mousedown", 4); break;
    case QEvent::MouseMove: ev("mousemove", 4); break;
    case QEvent::MouseButtonDblClick: dbl = true; ev("dblclick", 0); break;
    case QEvent::MouseButtonRelease: {
        bool inside = c && c->kind == K_OCX ? ocx_inside(c->ocx, cw, lp.toPoint()) : cw->rect().contains(lp.toPoint());
        if (c && c->kind == K_COMMAND) {        /* a CommandButton: Click comes before MouseUp (f_mouse) */
            QMouseEvent copy(me->type(), ev_pos(me), ev_scenepos(me), ev_gpos(me), me->button(), me->buttons(), me->modifiers());
            copy.setTimestamp(me->timestamp());
            QPointer<QObject> keep(obj);
            mouse_bypass++;
            QCoreApplication::sendEvent(obj, &copy);
            mouse_bypass--;
            if (c->form) ev("mouseup", 4);
            dbl = false;
            return true;
        }
        ev("mouseup", 4);
        /* Click: the left button; on a form any button (f_mouse: a right click) */
        if (own_click && !dbl && inside && (button == 1 || !c) && st->loaded) ev("click", 0);
        dbl = false;
        break;
    }
    default: break;
    }
    return false;
}

/* the keys a user types go through the same way */
static void key_track(QKeyEvent *ke, bool down);   /* GetAsyncKeyState (winapi_gui.inc) */
class KeyFilter : public QObject {
protected:
    bool eventFilter(QObject *obj, QEvent *e) override
    {
        QEvent::Type et = e->type();
        if (et == QEvent::MouseButtonPress || et == QEvent::MouseButtonRelease || et == QEvent::MouseButtonDblClick || et == QEvent::MouseMove)
            return mouse_event(obj, static_cast<QMouseEvent *>(e));
        if (et == QEvent::WhatsThis || et == QEvent::QueryWhatsThis) return help_whatsthis_event(obj, e);
        if (bypass || (et != QEvent::KeyPress && et != QEvent::KeyRelease)) return false;
        if (!static_cast<QKeyEvent *>(e)->isAutoRepeat()) key_track(static_cast<QKeyEvent *>(e), et == QEvent::KeyPress);
        if (et == QEvent::KeyPress && static_cast<QKeyEvent *>(e)->key() == Qt::Key_F1)   /* F1 in an open menu */
            if (auto *mn = qobject_cast<QMenu *>(obj)) if (help_menu_f1(mn)) return true;
        if (obj != QApplication::focusWidget()) return false;
        FormState *st = form_of_widget(obj);
        if (!st || !st->loaded) return false;
        auto *ke = static_cast<QKeyEvent *>(e);
        int vk = vk_of_qt(ke->key(), ke->text());
        if (!vk) return false;
        if (e->type() == QEvent::KeyPress) key_down(st, ke->key(), ke->modifiers(), ke->text(), vk, ke->isAutoRepeat());
        else if (!ke->isAutoRepeat()) key_up(st, ke->key(), ke->modifiers(), ke->text(), vk);
        return true;
    }
};

/* ---------------------------------------------------------------- SendKeys */
struct SKey { int key; int vk; QString text; };
static bool skey_named(const QString &name, SKey *k)
{
    static const std::map<QString, std::pair<int, int>> m = {
        { "BACKSPACE", { Qt::Key_Backspace, 8 } }, { "BS", { Qt::Key_Backspace, 8 } }, { "BKSP", { Qt::Key_Backspace, 8 } },
        { "BREAK", { Qt::Key_Pause, 19 } }, { "CAPSLOCK", { Qt::Key_CapsLock, 20 } }, { "DELETE", { Qt::Key_Delete, 46 } },
        { "DEL", { Qt::Key_Delete, 46 } }, { "DOWN", { Qt::Key_Down, 40 } }, { "END", { Qt::Key_End, 35 } },
        { "ENTER", { Qt::Key_Return, 13 } }, { "ESC", { Qt::Key_Escape, 27 } }, { "HELP", { Qt::Key_Help, 47 } },
        { "HOME", { Qt::Key_Home, 36 } }, { "INSERT", { Qt::Key_Insert, 45 } }, { "INS", { Qt::Key_Insert, 45 } },
        { "LEFT", { Qt::Key_Left, 37 } }, { "NUMLOCK", { Qt::Key_NumLock, 144 } }, { "PGDN", { Qt::Key_PageDown, 34 } },
        { "PGUP", { Qt::Key_PageUp, 33 } }, { "PRTSC", { Qt::Key_Print, 44 } }, { "RIGHT", { Qt::Key_Right, 39 } },
        { "SCROLLLOCK", { Qt::Key_ScrollLock, 145 } }, { "TAB", { Qt::Key_Tab, 9 } }, { "UP", { Qt::Key_Up, 38 } } };
    QString u = name.toUpper();
    auto it = m.find(u);
    if (it != m.end()) {
        k->key = it->second.first; k->vk = it->second.second;
        k->text = k->vk == 8 ? QString("\b") : k->vk == 13 ? QString("\r") : k->vk == 27 ? QString("\x1b") : k->vk == 9 ? QString("\t") : QString();
        return true;
    }
    if (u.size() >= 2 && u[0] == 'F') {
        bool ok; int f = u.mid(1).toInt(&ok);
        if (ok && f >= 1 && f <= 16) { k->key = Qt::Key_F1 + f - 1; k->vk = 111 + f; k->text.clear(); return true; }
    }
    return false;
}
static SKey skey_char(QChar ch, bool *shift)
{
    SKey k; k.text = QString(ch); *shift = false;
    ushort u = ch.unicode();
    if (u >= 'a' && u <= 'z') { k.key = Qt::Key_A + (u - 'a'); k.vk = 'A' + (u - 'a'); }
    else if (u >= 'A' && u <= 'Z') { k.key = Qt::Key_A + (u - 'A'); k.vk = u; *shift = true; }
    else if (u >= '0' && u <= '9') { k.key = Qt::Key_0 + (u - '0'); k.vk = u; }
    else if (u == ' ') { k.key = Qt::Key_Space; k.vk = 32; }
    else {
        k.key = Qt::Key_unknown;
        k.vk = vk_of_qt(Qt::Key_unknown, k.text);
        static const char *shifted = ":+<_>?~{|}\")!@#$%^&*(";
        if (ch.toLatin1() && strchr(shifted, ch.toLatin1())) *shift = true;
        if (!k.vk) k.vk = 0xE7;                 /* VK_PACKET: a character without a key */
    }
    return k;
}

static FormState *key_target_form()
{
    if (active_form && fstate(active_form)->loaded && fstate(active_form)->visible) return fstate(active_form);
    for (auto it = loaded_forms.rbegin(); it != loaded_forms.rend(); ++it) if (fstate(*it)->visible) return fstate(*it);
    return nullptr;
}

/* MsgBox, InputBox: the keys go to the dialog as they are */
static bool to_modal(int key, int mods, const QString &text, bool down, bool up)
{
    QWidget *m = QApplication::activePopupWidget();
    if (!m) m = QApplication::activeModalWidget();
    if (!m) return false;
    Qt::KeyboardModifiers qm;
    if (mods & 1) qm |= Qt::ShiftModifier;
    if (mods & 2) qm |= Qt::ControlModifier;
    if (mods & 4) qm |= Qt::AltModifier;
    QWidget *t = m->focusWidget() ? m->focusWidget() : m;
    if (down) deliver(t, QEvent::KeyPress, key, qm, text);
    t = QApplication::activeModalWidget() == m && m->focusWidget() ? m->focusWidget() : nullptr;
    if (up && t) deliver(t, QEvent::KeyRelease, key, qm, text);
    return true;
}

static void send_one(const SKey &k, int mods)
{
    if (to_modal(k.key, mods, k.text, true, true)) return;
    FormState *st = key_target_form();
    if (!st) return;
    Qt::KeyboardModifiers m;
    if (mods & 1) m |= Qt::ShiftModifier;
    if (mods & 2) m |= Qt::ControlModifier;
    if (mods & 4) m |= Qt::AltModifier;
    key_down(st, k.key, m, k.text, k.vk);
    if (vb_errpend) return;
    st = key_target_form();
    if (st) key_up(st, k.key, m, k.text, k.vk);
}
static void send_mod(int which, bool down, int mods)
{
    static const int keys[] = { Qt::Key_Shift, Qt::Key_Control, Qt::Key_Alt }, vks[] = { 16, 17, 18 };
    if (to_modal(keys[which], mods, QString(), down, !down)) return;
    FormState *st = key_target_form();
    if (!st) return;
    Qt::KeyboardModifiers m;
    if (mods & 1) m |= Qt::ShiftModifier;
    if (mods & 2) m |= Qt::ControlModifier;
    if (mods & 4) m |= Qt::AltModifier;
    if (down) key_down(st, keys[which], m, QString(), vks[which]);
    else key_up(st, keys[which], m, QString(), vks[which]);
}

static void sendkeys_run(const QString &s)
{
    int i = 0, n = (int)s.size();
    std::function<bool(int)> one;             /* sends one item with the modifiers already down */
    auto with_mods = [&](const SKey &k, int mods, bool shift) {
        int m = mods | (shift ? 1 : 0);
        bool extra = shift && !(mods & 1);
        if (extra) send_mod(0, true, mods | 1);
        send_one(k, m);
        if (extra) send_mod(0, false, mods);
    };
    std::function<bool(int)> item = [&](int mods) -> bool {
        if (i >= n || vb_errpend) return false;
        QChar ch = s[i];
        if (ch == '+' || ch == '^' || ch == '%') {
            int bit = ch == '+' ? 1 : ch == '^' ? 2 : 4, which = ch == '+' ? 0 : ch == '^' ? 1 : 2;
            i++;
            send_mod(which, true, mods | bit);
            if (i < n && s[i] == '(') {
                i++;
                while (i < n && s[i] != ')') if (!item(mods | bit)) break;
                if (i < n) i++;
            } else item(mods | bit);
            send_mod(which, false, mods);
            return true;
        }
        if (ch == '~') { SKey k; skey_named("ENTER", &k); i++; send_one(k, mods); return true; }
        if (ch == '{') {
            int close = s.indexOf('}', i + 2);
            if (close < 0) { vb_raise(5); return false; }
            QString body = s.mid(i + 1, close - i - 1);
            i = close + 1;
            int count = 1;
            int sp = body.lastIndexOf(' ');
            if (sp > 0) { bool ok; int c = body.mid(sp + 1).toInt(&ok); if (ok) { count = c; body = body.left(sp); } }
            SKey k; bool shift = false;
            if (body.size() == 1) k = skey_char(body[0], &shift);
            else if (!skey_named(body, &k)) { vb_raise(5); return false; }
            for (int r = 0; r < count && !vb_errpend; r++) with_mods(k, mods, shift);
            return true;
        }
        if (ch == '}' || ch == '(' || ch == ')') { vb_raise(5); return false; }
        bool shift; SKey k = skey_char(ch, &shift);
        i++;
        with_mods(k, mods, shift);
        return true;
    };
    while (i < n && !vb_errpend) if (!item(0)) break;
}

static void sendkeys_hook(vbs keys, vb_i32 wait)
{
    ensure_app();
    QString s = qs(keys);
    flush_pending();                            /* what the message loop had to do first (GotFocus ...) */
    if (wait) { sendkeys_run(s); flush_pending(); }
    else QTimer::singleShot(0, [s]() { sendkeys_run(s); });
}


/* ---------------------------------------------------------------- MsgBox, InputBox */
static QWidget *dialog_parent()
{
    if (active_form && fstate(active_form)->loaded && fstate(active_form)->visible) return fstate(active_form)->w;
    return nullptr;
}
static bool greek() { return vb_lcid == 1032; }

class VBDialog : public QDialog {
public:
    int esc = 0;                                /* what Esc and the close box give; 0: nothing */
    int result_code = 0;
    bool help = false;                          /* MsgBox with a help file: F1 opens it (help.inc) */
    /* no ? in the title bar (Qt 5 gives dialogs one on Windows; VB6's MsgBox and InputBox: only X) */
    explicit VBDialog(QWidget *parent = nullptr) : QDialog(parent) { setWindowFlag(Qt::WindowContextHelpButtonHint, false); }
protected:
    void reject() override { if (esc) { result_code = esc; QDialog::reject(); } }
    void keyPressEvent(QKeyEvent *e) override
    {
        if (help && e->key() == Qt::Key_F1 && !(e->modifiers() & (Qt::ControlModifier | Qt::AltModifier))) { msgbox_help(); return; }
        QDialog::keyPressEvent(e);
    }
};

static bool monkey_on;
static void monkey_log(const QString &what);
static vb_i32 msgbox_hook(vbs prompt, vb_i32 buttons, vbs title)
{
    ensure_app();
    if (monkey_on) {                            /* the test robot: the first button, at once */
        /* the first button; the last one when the same question comes again and again (Retry) */
        static QString last; static int same;
        QString p = qs(prompt);
        same = p == last ? same + 1 : 0; last = p;
        if (same < 3 || same % 50 == 0) monkey_log("MsgBox: " + p.replace('\n', ' ').left(120));
        static const vb_i32 first[] = { 1, 1, 3, 6, 6, 4 }, final[] = { 1, 2, 5, 2, 7, 2 };
        int k = buttons & 7; if (k > 5) k = 5;
        return same >= 2 ? final[k] : first[k];
    }
    flush_pending();
    VBDialog dlg(dialog_parent());
    dlg.setWindowTitle(qs(title));
    dlg.setModal(true);
    auto *outer = new QVBoxLayout(&dlg);
    auto *top = new QHBoxLayout;
    outer->addLayout(top);
    QStyle::StandardPixmap icons[] = { QStyle::SP_MessageBoxCritical, QStyle::SP_MessageBoxQuestion, QStyle::SP_MessageBoxWarning, QStyle::SP_MessageBoxInformation };
    int ic = (buttons >> 4) & 7;
    if (ic >= 1 && ic <= 4) {
        auto *il = new QLabel(&dlg);
        il->setPixmap(dlg.style()->standardIcon(icons[ic - 1]).pixmap(32, 32));
        il->setAlignment(Qt::AlignTop);
        top->addWidget(il);
    }
    QString text = qs(prompt); text.replace("\r\n", "\n"); text.replace('\r', '\n');
    auto *tl = new QLabel(text, &dlg);
    tl->setTextFormat(Qt::PlainText);
    top->addWidget(tl, 1);
    struct B { const char *en, *el; int code; };
    static const B ok = { "OK", "OK", 1 }, cancel = { "Cancel", "Άκυρο", 2 }, abort_ = { "&Abort", "&Ματαίωση", 3 },
                   retry = { "&Retry", "&Επανάληψη", 4 }, ignore = { "&Ignore", "&Παράβλεψη", 5 }, yes = { "&Yes", "&Ναι", 6 }, no = { "&No", "Ό&χι", 7 };
    std::vector<B> set;
    switch (buttons & 7) {
    case 1: set.push_back(ok); set.push_back(cancel); break;
    case 2: set.push_back(abort_); set.push_back(retry); set.push_back(ignore); break;
    case 3: set.push_back(yes); set.push_back(no); set.push_back(cancel); break;
    case 4: set.push_back(yes); set.push_back(no); break;
    case 5: set.push_back(retry); set.push_back(cancel); break;
    default: set.push_back(ok); break;
    }
    for (auto &b : set) if (b.code == 2 || (set.size() == 1 && b.code == 1)) dlg.esc = b.code;
    auto *row = new QHBoxLayout;
    row->addStretch();
    outer->addLayout(row);
    int def = (buttons >> 8) & 3;
    if (def >= (int)set.size() + ((buttons & 0x4000) ? 1 : 0)) def = 0;   /* the Help button can be the default */
    QPushButton *defb = nullptr;
    for (size_t i = 0; i < set.size(); i++) {
        auto *pb = new QPushButton(QString::fromUtf8(greek() ? set[i].el : set[i].en), &dlg);
        pb->setAutoDefault(true);
        int code = set[i].code;
        VBDialog *d = &dlg;
        QObject::connect(pb, &QPushButton::clicked, [d, code]() { d->result_code = code; d->accept(); });
        row->addWidget(pb);
        if ((int)i == def) { pb->setDefault(true); defb = pb; }
    }
    {                                           /* vbMsgBoxHelpButton: Help opens the help, the MsgBox stays (h1) */
        vb_i32 mk = vb_mark();
        vbs hf;
        vb_msgbox_helpinfo(&hf);
        dlg.help = hf.len > 0;
        vb_release(mk);
    }
    if (buttons & 0x4000) {
        auto *pb = new QPushButton(QString::fromUtf8(greek() ? "Βοήθεια" : "Help"), &dlg);
        pb->setAutoDefault(true);
        QObject::connect(pb, &QPushButton::clicked, []() { msgbox_help(); });
        row->addWidget(pb);
        if (def == (int)set.size() && !defb) { pb->setDefault(true); defb = pb; }
    }
    row->addStretch();
    if (defb) defb->setFocus();
    dlg.exec();
    return dlg.result_code ? dlg.result_code : (dlg.esc ? dlg.esc : 1);
}

static vbs inputbox_hook(vbs prompt, vbs title, vbs def)
{
    if (monkey_on) { monkey_log("InputBox: " + qs(prompt).left(80)); return vb_scopy(def); }
    ensure_app();
    flush_pending();
    VBDialog dlg(dialog_parent());
    dlg.setWindowTitle(qs(title));
    dlg.setModal(true);
    dlg.esc = 2;
    auto *grid = new QGridLayout(&dlg);
    QString text = qs(prompt); text.replace("\r\n", "\n");
    auto *tl = new QLabel(text, &dlg);
    tl->setTextFormat(Qt::PlainText); tl->setWordWrap(true); tl->setAlignment(Qt::AlignTop | Qt::AlignLeft);
    grid->addWidget(tl, 0, 0, 2, 1);
    auto *okb = new QPushButton("OK", &dlg);
    auto *cb = new QPushButton(QString::fromUtf8(greek() ? "Άκυρο" : "Cancel"), &dlg);
    okb->setAutoDefault(true); cb->setAutoDefault(true); okb->setDefault(true);
    grid->addWidget(okb, 0, 1); grid->addWidget(cb, 1, 1);
    auto *le = new QLineEdit(qs(def), &dlg);
    grid->addWidget(le, 2, 0, 1, 2);
    VBDialog *d = &dlg;
    QObject::connect(okb, &QPushButton::clicked, [d]() { d->result_code = 1; d->accept(); });
    QObject::connect(cb, &QPushButton::clicked, [d]() { d->result_code = 2; d->accept(); });
    le->selectAll();
    le->setFocus();
    dlg.exec();
    return dlg.result_code == 1 ? vs(le->text()) : vb_stmp(0);
}

/* ---------------------------------------------------------------- menus */
static QKeySequence shortcut_of(const QString &raw)
{
    QString s = raw.trimmed(), out;
    int i = 0;
    while (i < s.size() && (s[i] == '^' || s[i] == '+' || s[i] == '%')) {
        out += s[i] == '^' ? "Ctrl+" : s[i] == '+' ? "Shift+" : "Alt+";
        i++;
    }
    QString key = s.mid(i);
    if (key.startsWith('{') && key.endsWith('}')) {
        key = key.mid(1, key.size() - 2).toUpper();
        if (key == "DEL") key = "Del"; else if (key == "INSERT") key = "Ins"; else if (key == "BKSP") key = "Backspace";
    }
    return QKeySequence(out + key);
}

static void menu_apply(Ctl *c)
{
    QAction *a = c->act;
    if (!a) return;
    a->setText(a->isSeparator() ? QString() : c->caption);    /* a separator's "-" is not drawn (Fusion shows a text) */
    a->setEnabled(c->enabled);
    a->setVisible(c->visible);
    a->setCheckable(c->checked || a->isCheckable());
    a->setChecked(c->checked);
}

static void build_menus(FormState *st)
{
    const vbf_formdef *d = st->def;
    std::vector<Ctl *> bydef(d->nctls, nullptr);
    for (vbobj o : st->ctls) { Ctl *c = C(o); for (int i = 0; i < d->nctls; i++) if (&d->ctls[i] == c->def && c->index == d->ctls[i].index) bydef[i] = c; }
    bool any = false;
    for (int i = 0; i < d->nctls; i++) if (bydef[i] && bydef[i]->kind == K_MENU) any = true;
    if (!any) return;
    st->menubar = new FormMenuBar(st->w);
    st->w->layout()->setMenuBar(st->menubar);
    for (int i = 0; i < d->nctls; i++) {
        Ctl *c = bydef[i];
        if (!c || c->kind != K_MENU) continue;
        bool has_children = false;
        for (int j = 0; j < d->nctls; j++) if (d->ctls[j].parent == i) has_children = true;
        Ctl *parent = d->ctls[i].parent >= 0 ? bydef[d->ctls[i].parent] : nullptr;
        vbobj self = c->self;
        if (has_children) {
            c->menu = parent && parent->menu ? parent->menu->addMenu(c->caption) : st->menubar->addMenu(c->caption);
            c->act = c->menu->menuAction();
            QObject::connect(c->menu, &QMenu::aboutToShow, [self]() { fire_ctl(C(self), "click"); });
        } else if (c->caption == "-") {
            c->act = parent && parent->menu ? parent->menu->addSeparator() : st->menubar->addSeparator();
        } else {
            c->act = parent && parent->menu ? parent->menu->addAction(c->caption) : st->menubar->addAction(c->caption);
            QObject::connect(c->act, &QAction::triggered, [self]() { fire_ctl(C(self), "click"); });
            if (!c->shortcut.isEmpty()) { c->act->setShortcut(shortcut_of(c->shortcut)); c->act->setShortcutContext(Qt::WindowShortcut); }
        }
        if (c->checked) c->act->setCheckable(true);
        c->act->setProperty("vbmenu", QVariant::fromValue((qulonglong)(quintptr)self));   /* F1 in a menu (help.inc) */
        menu_apply(c);
    }
}


/* ---------------------------------------------------------------- Picture properties */
static int border_px(Ctl *c) { return c->kind == K_PICTURE && propn(c->def->props, "BorderStyle", 1) != 0 ? 2 : 0; }
static void image_show(Ctl *c)
{
    auto *l = qobject_cast<QLabel *>(c->w.data());
    if (!l) return;
    Pic *p = pic_of(c->picture);
    l->setScaledContents(c->stretch);
    l->setPixmap(p && p->type ? QPixmap::fromImage(p->img) : QPixmap());
}
static void set_picture(Ctl *c, vbobj pic)
{
    vbobj np = pic && pic_of(pic) ? vb_oid(pic) : new_pic(QImage(), 0);
    vb_oaddref(np);
    if (c->picture) vb_orelease(c->picture);
    c->picture = np;
    Pic *p = pic_of(np);
    if (c->kind == K_IMAGE) {
        if (!c->stretch && p->type) { c->width = p->img.width() * TPP; c->height = p->img.height() * TPP; ctl_geometry(c); }
        image_show(c);
    } else if (c->g) {
        if (c->autosize && p->type) {
            int b = border_px(c);
            c->width = (p->img.width() + 2 * b) * TPP; c->height = (p->img.height() + 2 * b) * TPP;
            ctl_geometry(c);
        }
        c->g->pic = p->img;
        gfx_clear(c->g);
    }
}
static vbobj picture_get(vbobj *slot)
{
    if (!*slot) { *slot = new_pic(QImage(), 0); vb_oaddref(*slot); }
    return *slot;
}

/* Left, Top, Width, Height of a control are in the ScaleMode of its container (the form, a
 * PictureBox; a Frame has twips); the .frm and the runtime keep twips */
static Gfx *coord_gfx(Ctl *c)
{
    if (c->container) { Ctl *k = C(c->container); return k && k->g ? k->g : nullptr; }
    return c->form ? &c->form->g : nullptr;
}
static double to_twips(Ctl *c, double v, bool vert, bool pos)
{
    Gfx *g = coord_gfx(c);
    if (!g || g->scalemode == 1) return v;
    if (g->scalemode != 0) return v / units_per_twip(g->scalemode, vert);
    double span = vert ? g->sh : g->sw, tw = (vert ? g->ph : g->pw) * g->tpp;
    if (!span || !tw) return v;
    return ((pos ? v - (vert ? g->st : g->sl) : v) * tw / span);
}
static double from_twips(Ctl *c, double t, bool vert, bool pos)
{
    Gfx *g = coord_gfx(c);
    if (!g || g->scalemode == 1) return t;
    if (g->scalemode != 0) return t * units_per_twip(g->scalemode, vert);
    double span = vert ? g->sh : g->sw, tw = (vert ? g->ph : g->pw) * g->tpp;
    if (!span || !tw) return t;
    return t * span / tw + (pos ? (vert ? g->st : g->sl) : 0);
}

/* the Picture of a CommandButton, CheckBox or OptionButton: shown when its Style is 1 - Graphical */
static void button_icon(Ctl *c)
{
    auto *b = qobject_cast<QAbstractButton *>(c->w.data());
    if (!b) return;
    vbobj p = c->btnpic[c->enabled || !c->btnpic[2] ? 0 : 2];
    Pic *pp = p ? pic_of(p) : nullptr;
    if (c->def && propn(c->def->props, "Style", 0) == 1 && pp && !pp->img.isNull()) {
        QPixmap px = QPixmap::fromImage(pp->img);
        b->setIcon(QIcon(px));
        b->setIconSize(px.size());
    } else {
        b->setIcon(QIcon());
    }
}

/* members of Image, Shape, Line, the scroll bars and the file controls; *found = 0: not one */
static vbvar more_member(Ctl *c, const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    *found = 1;
    if (n == "picture" && (c->kind == K_IMAGE || c->kind == K_PICTURE)) {
        if (set) { set_picture(c, val.vt == VT_DISPATCH ? val.u.o : nullptr); return vb_v_empty(); }
        return vb_vtmp(vb_v_obj(picture_get(&c->picture)));
    }
    switch (c->kind) {
    case K_IMAGE:
        if (n == "stretch") { if (!set) return vb_v_bool(c->stretch ? -1 : 0); c->stretch = vb_v2bool(val) != 0; image_show(c); return vb_v_empty(); }
        if (n == "borderstyle") return vb_v_i((vb_i16)propn(c->def->props, "BorderStyle", 0));
        if (n == "backcolor" || n == "forecolor" || n == "hwnd" || n == "tabindex" || n == "setfocus" || n.compare(0, 4, "font") == 0) { *found = 0; return vb_v_empty(); }
        break;
    case K_PICTURE:
        if (n == "autosize") {
            if (!set) return vb_v_bool(c->autosize ? -1 : 0);
            c->autosize = vb_v2bool(val) != 0;
            if (c->autosize) set_picture(c, c->picture);
            return vb_v_empty();
        }
        break;
    case K_SHAPE: case K_LINE: {
        int *ip = n == "borderstyle" ? &c->borderstyle : n == "borderwidth" ? &c->borderwidth : nullptr;
        if (c->kind == K_SHAPE && !ip) ip = n == "shape" ? &c->shape : n == "backstyle" ? &c->backstyle : n == "fillstyle" ? &c->fillstyle : nullptr;
        if (ip) {
            if (!set) return vb_v_i((vb_i16)*ip);
            int v = vb_v2i(val);
            if (v < 0 || (n == "borderwidth" && (v < 1 || v > 8192)) || (n == "shape" && v > 5) || (n == "fillstyle" && v > 7) || (n == "borderstyle" && v > 6) || (n == "backstyle" && v > 1)) { vb_raise(380); return vb_v_empty(); }
            *ip = v; ctl_geometry(c); c->w->update();
            return vb_v_empty();
        }
        long *lp = n == "bordercolor" ? &c->bordercolor : (c->kind == K_SHAPE && n == "fillcolor") ? &c->fillcolor : (c->kind == K_SHAPE && n == "backcolor") ? &c->back : nullptr;
        if (lp) { if (!set) return vb_v_l((vb_i32)*lp); *lp = vb_v2l(val); c->w->update(); return vb_v_empty(); }
        if (n == "drawmode") return vb_v_i(13);
        if (c->kind == K_LINE) {
            double *dp = n == "x1" ? &c->x1 : n == "y1" ? &c->y1 : n == "x2" ? &c->x2 : n == "y2" ? &c->y2 : nullptr;
            bool vert = n == "y1" || n == "y2";
            if (dp) { if (!set) return vb_v_f((float)from_twips(c, *dp, vert, true)); *dp = to_twips(c, vb_v2d(val), vert, true); line_geometry(c); return vb_v_empty(); }
            if (n == "left" || n == "top" || n == "width" || n == "height" || n == "move") { *found = 0; return vb_v_empty(); }
        }
        if (n == "visible" || n == "left" || n == "top" || n == "width" || n == "height" || n == "move" || n == "container" || n == "zorder" || n == "refresh") break;
        *found = 0;
        return vb_v_empty();
    }
    case K_HSCROLL: case K_VSCROLL: {
        if (n == "value") { if (!set) return vb_v_i((vb_i16)c->sval); scroll_set(c, vb_v2i(val)); return vb_v_empty(); }
        int *ip = n == "min" ? &c->smin : n == "max" ? &c->smax : n == "smallchange" ? &c->small : n == "largechange" ? &c->large : nullptr;
        if (ip) {
            if (!set) return vb_v_i((vb_i16)*ip);
            int v = vb_v2i(val);
            if ((n == "smallchange" || n == "largechange") && (v < 1)) { vb_raise(380); return vb_v_empty(); }
            *ip = v;
            int lo = std::min(c->smin, c->smax), hi = std::max(c->smin, c->smax), old = c->sval;
            c->sval = std::max(lo, std::min(hi, c->sval));
            scroll_sync(c);
            if (c->sval != old) fire_ctl(c, "change");
            return vb_v_empty();
        }
        break;
    }
    case K_DRIVE:
        if (n == "drive") {
            if (!set) return vstr(c->path);
            QString d = qs(vb_v2s(val)).trimmed();
            auto *cb = CB(c);
            int found_i = -1;
            for (int i = 0; cb && i < cb->count(); i++) if (!d.isEmpty() && cb->itemText(i).startsWith(d, Qt::CaseInsensitive)) { found_i = i; break; }
            if (found_i < 0) { vb_raise(68); return vb_v_empty(); }   /* Device unavailable */
            if (cb->itemText(found_i) != c->path) { QSignalBlocker b(cb); cb->setCurrentIndex(found_i); c->path = cb->itemText(found_i); fire_ctl(c, "change"); }
            return vb_v_empty();
        }
        if (n == "list" || n == "listcount" || n == "listindex") break;
        break;
    case K_DIR:
        if (n == "path") { if (!set) return vstr(shown_path(c->path)); dir_setpath(c, qs(vb_v2s(val))); return vb_v_empty(); }
        if (n == "listcount") return vb_v_i((vb_i16)dir_subs(c).size());
        if (n == "list") {
            int i = (int)arg_l(argc, argv, 0);
            if (set) { vb_raise(383); return vb_v_empty(); }
            if (i >= 0) { QStringList subs = dir_subs(c); return vstr(i < subs.size() ? shown_path(QDir(c->path).absoluteFilePath(subs[i])) : QString()); }
            QStringList par = dir_parents(c);
            int k = par.size() + i;             /* -1: the folder itself, -2 its parent ... */
            return vstr(k >= 0 ? shown_path(par[k]) : QString());
        }
        if (n == "listindex") {
            auto *lw = qobject_cast<QListWidget *>(c->w.data());
            int np = dir_parents(c).size();
            if (!set) return vb_v_i((vb_i16)(lw ? lw->currentRow() - np : -1));
            int i = vb_v2i(val);
            if (lw) { QSignalBlocker b(lw); lw->setCurrentRow(np + i); }
            fire_ctl(c, "click");
            return vb_v_empty();
        }
        break;
    case K_FILE:
        if (n == "path") {
            if (!set) return vstr(shown_path(c->path));
            QFileInfo fi(QDir(c->path).absoluteFilePath(native(qs(vb_v2s(val)))));
            if (!fi.exists() || !fi.isDir()) { vb_raise(76); return vb_v_empty(); }
            QString np = QDir::cleanPath(fi.absoluteFilePath());
            if (np != c->path) { c->path = np; file_fill(c); fire_ctl(c, "pathchange"); }
            return vb_v_empty();
        }
        if (n == "pattern") {
            if (!set) return vstr(c->pattern);
            QString np = qs(vb_v2s(val));
            if (np != c->pattern) { c->pattern = np; file_fill(c); fire_ctl(c, "patternchange"); }
            return vb_v_empty();
        }
        if (n == "filename") {
            auto *lw = LW(c);
            if (!set) return vstr(lw && lw->currentRow() >= 0 ? lw->currentItem()->text() : QString());
            QString v = native(qs(vb_v2s(val)));
            if (v.contains('*') || v.contains('?')) { vbvar a = vstr(v); return more_member(c, "pattern", VBM_LET, true, a, 1, &a, found); }
            QFileInfo fi(QDir(c->path).absoluteFilePath(v));
            if (fi.isDir()) { vbvar a = vstr(fi.absoluteFilePath()); return more_member(c, "path", VBM_LET, true, a, 1, &a, found); }
            return vb_v_empty();
        }
        {
            bool *bp = n == "archive" ? &c->f_archive : n == "hidden" ? &c->f_hidden : n == "normal" ? &c->f_normal : n == "readonly" ? &c->f_readonly : n == "system" ? &c->f_system : nullptr;
            if (bp) { if (!set) return vb_v_bool(*bp ? -1 : 0); *bp = vb_v2bool(val) != 0; file_fill(c); return vb_v_empty(); }
        }
        break;
    default: break;
    }
    (void)kind;
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- controls: members */
static vbvar ctl_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static void ctl_clear(vbobj o)
{
    Ctl *c = (Ctl *)o->ext;
    if (!c) return;
    delete c->timer;
    if (c->uc) uctl_release(c);
    if (c->bag) vb_orelease(c->bag);
    if (c->w) { c->w->hide(); c->w->setParent(nullptr); c->w->deleteLater(); }   /* (no event of it after c->g goes) */
    if (c->picture) vb_orelease(c->picture);
    delete c->g;
    bound_release(c);
    if (c->dfmt) vb_orelease(c->dfmt);
    ocx_delete(c->ocx);
    delete c;
    o->ext = nullptr;
}

#define CTL_CLASS(var, name, def) \
    static const vb_class var = { name, (vb_i32)sizeof(vbobjh), nullptr, nullptr, ctl_clear, nullptr, 0, def, nullptr, nullptr, nullptr, ctl_member, nullptr }
CTL_CLASS(CLS_COMMAND, "CommandButton", "value");
CTL_CLASS(CLS_LABEL, "Label", "caption");
CTL_CLASS(CLS_TEXT, "TextBox", "text");
CTL_CLASS(CLS_CHECK, "CheckBox", "value");
CTL_CLASS(CLS_OPTION, "OptionButton", "value");
CTL_CLASS(CLS_LIST, "ListBox", "text");
CTL_CLASS(CLS_COMBO, "ComboBox", "text");
CTL_CLASS(CLS_TIMER, "Timer", "enabled");
CTL_CLASS(CLS_FRAME, "Frame", "caption");
CTL_CLASS(CLS_PICTURE, "PictureBox", "picture");
CTL_CLASS(CLS_MENU, "Menu", "enabled");      /* the default properties of VB6 (Programmer's Guide) */
CTL_CLASS(CLS_IMAGE, "Image", "picture");
CTL_CLASS(CLS_SHAPE, "Shape", "shape");
CTL_CLASS(CLS_LINE, "Line", "visible");
CTL_CLASS(CLS_HSCROLL, "HScrollBar", "value");
CTL_CLASS(CLS_VSCROLL, "VScrollBar", "value");
CTL_CLASS(CLS_DRIVE, "DriveListBox", "drive");
CTL_CLASS(CLS_DIR, "DirListBox", "path");
CTL_CLASS(CLS_FILE, "FileListBox", "filename");

static const vb_class *ctl_class(Kind k, const std::string &type)
{
    switch (k) {
    case K_COMMAND: return &CLS_COMMAND; case K_LABEL: return &CLS_LABEL; case K_TEXT: return &CLS_TEXT;
    case K_CHECK: return &CLS_CHECK; case K_OPTION: return &CLS_OPTION; case K_LIST: return &CLS_LIST;
    case K_COMBO: return &CLS_COMBO; case K_TIMER: return &CLS_TIMER; case K_FRAME: return &CLS_FRAME;
    case K_PICTURE: return &CLS_PICTURE; case K_MENU: return &CLS_MENU;
    case K_IMAGE: return &CLS_IMAGE; case K_SHAPE: return &CLS_SHAPE; case K_LINE: return &CLS_LINE;
    case K_HSCROLL: return &CLS_HSCROLL; case K_VSCROLL: return &CLS_VSCROLL; case K_DRIVE: return &CLS_DRIVE;
    case K_DIR: return &CLS_DIR; case K_FILE: return &CLS_FILE;
    default: {
        static std::map<std::string, vb_class *> others;
        auto it = others.find(type);
        if (it != others.end()) return it->second;
        vb_class *c = new vb_class(CLS_PICTURE);
        c->name = strdup(type.c_str()); c->defmember = k == K_OCX ? ocx_default(type) : nullptr;
        if (k == K_USER) { const vb_class *u = uctl_class(type.c_str()); c->defmember = u ? u->defmember : nullptr; }
        if (k == K_DATA) c->defmember = "caption";
        others[type] = c;
        return c;
    }
    }
}

static int list_count(Ctl *c) { return LW(c) ? LW(c)->count() : CB(c) ? CB(c)->count() : 0; }
static int list_index(Ctl *c) { return LW(c) ? LW(c)->currentRow() : CB(c) ? CB(c)->currentIndex() : -1; }

static vbvar list_member(Ctl *c, const std::string &n, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    bool set = kind == VBM_LET || kind == VBM_SET;
    QListWidget *lw = LW(c); QComboBox *cb = CB(c);
    *found = 1;
    if (n == "additem") {
        QString item = arg_s(argc, argv, 0);
        int cnt = list_count(c), at = (int)arg_l(argc, argv, 1, cnt);
        if (at < 0 || at > cnt) { vb_raise(5); return vb_v_empty(); }
        if (lw) lw->insertItem(at, item);
        else {
            QSignalBlocker b(cb); QSignalBlocker b2(cb->lineEdit());
            QString t = cb->currentText();
            int was = cb->currentIndex();
            cb->insertItem(at, item);
            if (cb->isEditable()) cb->setEditText(t);
            else if (was < 0) cb->setCurrentIndex(-1);   /* a Dropdown List: no item chosen yet (ListIndex -1) */
        }
        c->itemdata.insert(c->itemdata.begin() + at, 0);
        c->newindex = at;
        return vb_v_empty();
    }
    if (n == "removeitem") {
        int at = (int)arg_l(argc, argv, 0);
        if (at < 0 || at >= list_count(c)) { vb_raise(5); return vb_v_empty(); }
        if (lw) delete lw->takeItem(at); else { QSignalBlocker b(cb); QSignalBlocker b2(cb->lineEdit()); cb->removeItem(at); }
        c->itemdata.erase(c->itemdata.begin() + at);
        return vb_v_empty();
    }
    if (n == "clear") {
        if (lw) lw->clear(); else { QSignalBlocker b(cb); QSignalBlocker b2(cb->lineEdit()); cb->clear(); }
        c->itemdata.clear();
        return vb_v_empty();
    }
    if (n == "listcount") return vb_v_i((vb_i16)list_count(c));
    if (n == "newindex") return vb_v_i((vb_i16)c->newindex);
    if (n == "list") {
        int i = (int)arg_l(argc, argv, 0);
        if (set) {
            if (i < 0 || i >= list_count(c)) { vb_raise(381); return vb_v_empty(); }
            QString v = arg_s(argc, argv, argc - 1);
            if (lw) lw->item(i)->setText(v); else cb->setItemText(i, v);
            return vb_v_empty();
        }
        if (i < 0 || i >= list_count(c)) return vstr(QString());
        return vstr(lw ? lw->item(i)->text() : cb->itemText(i));
    }
    if (n == "itemdata") {
        int i = (int)arg_l(argc, argv, 0);
        if (i < 0 || i >= (int)c->itemdata.size()) { vb_raise(381); return vb_v_empty(); }
        if (set) { c->itemdata[i] = arg_l(argc, argv, argc - 1); return vb_v_empty(); }
        return vb_v_l((vb_i32)c->itemdata[i]);
    }
    if (n == "listindex") {
        if (!set) return vb_v_i((vb_i16)list_index(c));
        int i = (int)arg_l(argc, argv, 0);
        if (i < -1 || i >= list_count(c)) { vb_raise(380); return vb_v_empty(); }
        if (i == list_index(c)) return vb_v_empty();
        if (lw) { QSignalBlocker b(lw); lw->setCurrentRow(i); }
        else { QSignalBlocker b(cb); QSignalBlocker b2(cb->lineEdit()); cb->setCurrentIndex(i); if (i < 0) cb->setEditText(QString()); }
        fire_ctl(c, "click");
        return vb_v_empty();
    }
    if (n == "selected" && lw) {
        int i = (int)arg_l(argc, argv, 0);
        if (i < 0 || i >= lw->count()) { vb_raise(381); return vb_v_empty(); }
        if (set) { lw->item(i)->setSelected(vb_v2bool(argv[argc - 1]) != 0); return vb_v_empty(); }
        return vb_v_bool(lw->item(i)->isSelected() ? -1 : 0);
    }
    if (n == "text") {
        if (set) { set_text(c, arg_s(argc, argv, 0), true); return vb_v_empty(); }
        return vstr(get_text(c));
    }
    *found = 0;
    return vb_v_empty();
}

static vbvar menu_member(Ctl *c, const std::string &n, bool set, vbvar val, int *found)
{
    *found = 1;
    if (n == "caption") { if (!set) return vstr(c->caption); c->caption = qs(vb_v2s(val)); }
    else if (n == "checked") { if (!set) return vb_v_bool(c->checked ? -1 : 0); c->checked = vb_v2bool(val) != 0; if (c->act) c->act->setCheckable(true); }
    else if (n == "enabled") { if (!set) return vb_v_bool(c->enabled ? -1 : 0); c->enabled = vb_v2bool(val) != 0; }
    else if (n == "visible") { if (!set) return vb_v_bool(c->visible ? -1 : 0); c->visible = vb_v2bool(val) != 0; }
    else if (n == "windowlist") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_bool(propn(c->def->props, "WindowList", 0) ? -1 : 0); }
    else if (n == "negotiateposition") return vb_v_i((vb_i16)propn(c->def->props, "NegotiatePosition", 0));
    else if (n == "helpcontextid") { if (!set) return vb_v_l(ctl_helpid(c, "helpcontextid")); c->misc["helpcontextid"] = vb_v2l(val); return vb_v_empty(); }
    else { *found = 0; return vb_v_empty(); }
    menu_apply(c);
    return vb_v_empty();
}

/* "Text (*.txt)|*.txt|All|*.*" -> the filters of Qt: "Text (*.txt)", "All (*.*)" */
static QStringList dlg_filters(const QString &f)
{
    QStringList parts = f.split('|'), out;
    for (int i = 0; i + 1 < parts.size(); i += 2) {
        QString label = parts[i].trimmed(), pats = parts[i + 1];
        pats.replace(';', ' ');
        int k = label.lastIndexOf('(');
        if (label.endsWith(')') && k > 0) label = label.left(k).trimmed();
        out << label + " (" + pats.trimmed() + ")";
    }
    return out;
}
static bool dlg_cancel(Ctl *c)
{
    if (c->dlg->cancelerror) {
        vb_i32 m = vb_mark();
        vb_raise_full(32755, vs(QString("CommonDialog")), vs(QString("Cancel was selected.")));
        vb_release(m);
    }
    return false;
}
static QString win_path(const QString &p)
{
#ifdef _WIN32
    return QDir::toNativeSeparators(p);
#else
    return p;
#endif
}
static void dlg_file(Ctl *c, bool save)
{
    CDlg *d = c->dlg;
    if (monkey_on) { monkey_log(QString("CommonDialog %1: Cancel").arg(save ? "ShowSave" : "ShowOpen")); dlg_cancel(c); return; }
    QStringList filters = dlg_filters(d->filter);
    QString sel = filters.isEmpty() ? QString() : filters.value(std::max(1, d->filterindex) - 1, filters.first());
    QString start = d->filename;
    if (!d->initdir.isEmpty() && QFileInfo(start).isRelative()) start = QDir(native(d->initdir)).filePath(start);
    start = native(start);
    QString title = d->title.isEmpty() ? (save ? "Save As" : "Open") : d->title;
    QFileDialog::Options opt;
    if (!(d->flags & 0x2)) opt |= QFileDialog::DontConfirmOverwrite;       /* cdlOFNOverwritePrompt */
    QStringList files;
    if (save) {
        QString f = QFileDialog::getSaveFileName(dialog_parent(), title, start, filters.join(";;"), &sel, opt);
        if (!f.isEmpty()) {
            if (!d->defaultext.isEmpty() && QFileInfo(f).suffix().isEmpty()) f += "." + d->defaultext;
            files << f;
        }
    } else if (d->flags & 0x200) {                                       /* cdlOFNAllowMultiselect */
        files = QFileDialog::getOpenFileNames(dialog_parent(), title, start, filters.join(";;"), &sel, opt);
    } else {
        QString f = QFileDialog::getOpenFileName(dialog_parent(), title, start, filters.join(";;"), &sel, opt);
        if (!f.isEmpty()) files << f;
    }
    if (files.isEmpty()) { dlg_cancel(c); return; }
    int fi = filters.indexOf(sel);
    if (fi >= 0) d->filterindex = fi + 1;
    if (files.size() == 1) {
        d->filename = win_path(files[0]);
        d->filetitle = QFileInfo(files[0]).fileName();
    } else {                                    /* the folder, then the names: Chr(0) between them (Explorer) */
        QString sep = d->flags & 0x80000 ? QString(QChar(0)) : QString(" ");
        QString r = win_path(QFileInfo(files[0]).absolutePath());
        for (const QString &f : files) r += sep + QFileInfo(f).fileName();
        d->filename = r; d->filetitle.clear();
    }
}
static vbvar dialog_member(Ctl *c, const std::string &n, bool set, vbvar val, int *found)
{
    CDlg *d = c->dlg;
    auto str = [&](QString *f) { if (set) { *f = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(*f); };
    auto num = [&](int *f) { if (set) { *f = vb_v2i(val); return vb_v_empty(); } return vb_v_i((vb_i16)*f); };
    auto lng = [&](long *f) { if (set) { *f = vb_v2l(val); return vb_v_empty(); } return vb_v_l((vb_i32)*f); };
    auto flag = [&](bool *f) { if (set) { *f = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(*f ? -1 : 0); };
    if (n == "filename") return str(&d->filename);
    if (n == "filetitle") return str(&d->filetitle);
    if (n == "filter") return str(&d->filter);
    if (n == "dialogtitle") return str(&d->title);
    if (n == "initdir") return str(&d->initdir);
    if (n == "defaultext") return str(&d->defaultext);
    if (n == "fontname") return str(&d->fontname);
    if (n == "filterindex") return num(&d->filterindex);
    if (n == "copies") return num(&d->copies);
    if (n == "frompage") return num(&d->frompage);
    if (n == "topage") return num(&d->topage);
    if (n == "min") return num(&d->min);
    if (n == "max") return num(&d->max);
    if (n == "orientation") return num(&d->orientation);
    if (n == "flags") return lng(&d->flags);
    if (n == "color") return lng(&d->color);
    if (n == "cancelerror") return flag(&d->cancelerror);
    if (n == "fontbold") return flag(&d->bold);
    if (n == "fontitalic") return flag(&d->italic);
    if (n == "fontunderline") return flag(&d->underline);
    if (n == "fontstrikethru") return flag(&d->strike);
    if (n == "fontsize") { if (set) { d->fontsize = vb_v2d(val); return vb_v_empty(); } return vb_v_f((float)d->fontsize); }
    if (n == "helpfile") { if (set) { d->helpfile = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(d->helpfile); }
    if (n == "helpkey") { if (set) { d->helpkey = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(d->helpkey); }
    if (n == "helpcommand") { if (set) { d->helpcommand = vb_v2l(val); return vb_v_empty(); } return vb_v_i((vb_i16)d->helpcommand); }
    if (n == "helpcontext") { if (set) { d->helpcontext = vb_v2l(val); return vb_v_empty(); } return vb_v_l(d->helpcontext); }
    if (n == "maxfilesize" || n == "printerdefault" || n == "hdc")
        return set ? vb_v_empty() : vb_v_l(0);
    if (n == "showopen" || n == "showsave") { dlg_file(c, n == "showsave"); return vb_v_empty(); }
    if (n == "showcolor") {
        if (monkey_on) { monkey_log("CommonDialog ShowColor: Cancel"); dlg_cancel(c); return vb_v_empty(); }
        QColor q = QColorDialog::getColor(qcolor(d->color), dialog_parent(), d->title.isEmpty() ? "Color" : d->title);
        if (!q.isValid()) { dlg_cancel(c); return vb_v_empty(); }
        d->color = q.red() | (q.green() << 8) | (q.blue() << 16);
        return vb_v_empty();
    }
    if (n == "showfont") {
        if (monkey_on) { monkey_log("CommonDialog ShowFont: Cancel"); dlg_cancel(c); return vb_v_empty(); }
        QFont f(d->fontname);
        f.setPointSizeF(d->fontsize >= 1 && d->fontsize < 2160 ? d->fontsize : 8.25);
        f.setBold(d->bold); f.setItalic(d->italic); f.setUnderline(d->underline); f.setStrikeOut(d->strike);
        bool ok = false;
        QFont r = QFontDialog::getFont(&ok, f, dialog_parent(), d->title.isEmpty() ? "Font" : d->title);
        if (!ok) { dlg_cancel(c); return vb_v_empty(); }
        d->fontname = r.family(); d->fontsize = r.pointSizeF();
        d->bold = r.bold(); d->italic = r.italic(); d->underline = r.underline(); d->strike = r.strikeOut();
        if (d->flags & 0x100) d->color = d->color;                       /* cdlCFEffects: the color stays */
        return vb_v_empty();
    }
    if (n == "showprinter") {                   /* the printer of the program is the PDF of Printer */
        if (monkey_on) { monkey_log("CommonDialog ShowPrinter"); return vb_v_empty(); }
        QMessageBox::information(dialog_parent(), d->title.isEmpty() ? "Print" : d->title,
                                 "The printout is written as a PDF file (Printer object).");
        return vb_v_empty();
    }
    if (n == "showhelp") { dlg_showhelp(c); return vb_v_empty(); }
    if (n == "aboutbox") return vb_v_empty();
    if (n == "action") {                        /* the VB3 way: Action = 1 Open, 2 Save, 3 Color, 4 Font, 5 Printer, 6 Help */
        if (!set) return vb_v_i(0);
        static const char *const act[] = { "", "showopen", "showsave", "showcolor", "showfont", "showprinter", "showhelp" };
        vb_i32 a = vb_v2l(val);
        if (a < 0 || a > 6) { vb_raise(380); return vb_v_empty(); }
        if (a == 0) return vb_v_empty();
        return dialog_member(c, act[a], false, vb_v_empty(), found);
    }
    *found = 0;
    return vb_v_empty();
}

/* DDE (LinkMode, LinkTopic ...): there is no DDE here. The properties are kept; a control that starts a
 * conversation (LinkMode 1, 2 or 3) gets error 282 "No foreign application responded to a DDE initiate",
 * as VB6 when no program answers - programs use it to find out that they are the first instance. */
static std::map<std::pair<const void *, std::string>, QString> dde_props;
static bool dde_member(const void *who, bool is_ctl, const std::string &n, bool set, vbvar val, vbvar *r)
{
    *r = vb_v_empty();
    if (n == "linkexecute" || n == "linkpoke" || n == "linkrequest" || n == "linksend") return true;
    bool num = n == "linkmode" || n == "linktimeout";
    if (!(num || n == "linktopic" || n == "linkitem")) return false;
    auto key = std::make_pair(who, n);
    if (set) {
        if (n == "linkmode" && is_ctl && vb_v2l(val) != 0) { vb_raise(282); return true; }
        dde_props[key] = num ? QString::number(vb_v2l(val)) : qs(vb_v2s(val));
        return true;
    }
    auto it = dde_props.find(key);
    if (num) *r = vb_v_i((vb_i16)(it != dde_props.end() ? it->second.toInt() : n == "linkmode" ? 0 : 50));
    else *r = vstr(it != dde_props.end() ? it->second : QString());
    return true;
}
static vbvar ctl_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    Ctl *c = C(o);
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    if (!c || !c->form) { vb_raise(91); return vb_v_empty(); }
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    if (n == "name") { if (set) { vb_raise(382); return vb_v_empty(); } return vstr(c->name); }
    if (n == "index") { if (c->index < 0) { vb_raise(343); return vb_v_empty(); } return vb_v_i((vb_i16)c->index); }
    if (n == "tag") { if (set) { c->tag = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(c->tag); }
    if (n == "parent") return vb_vtmp(vb_v_obj(c->form->form));
    if (c->kind == K_USER) {                    /* its own members, then those of the Extender */
        vbvar r = uctl_member(c, n, kind, set, val, argc, argv, found);
        if (*found || vb_errpend) return r;
        *found = 1;
    }
    if (c->kind == K_DIALOG) return dialog_member(c, n, set, val, found);
    if (c->kind == K_DATA) {
        vbvar r = data_member(c, n, kind, set, val, argc, argv, found);
        if (*found || vb_errpend) return r;
        *found = 1;
    } else if (n.compare(0, 4, "data") == 0 && !(c->kind == K_OCX && ocx_data_aware(c->ocx))) {   /* DataField, DataSource, DataChanged of a bound control */
        vbvar r = bound_member(c, n, set, val, found);
        if (*found || vb_errpend) return r;
        *found = 1;
    }
    if (c->kind == K_OCX && c->ocx) {
        vbvar r = ocx_member(c, n, kind, set, val, argc, argv, found);
        if (*found || vb_errpend) return r;
        *found = 1;
    }
    if (c->kind == K_MENU) {
        vbvar r = menu_member(c, n, set, val, found);
        return r;
    }
    if (c->kind >= K_IMAGE || c->kind == K_PICTURE) {
        vbvar r = more_member(c, n, kind, set, val, argc, argv, found);
        if (*found || vb_errpend) return r;
        *found = 1;
        if (c->kind == K_IMAGE && (n == "backcolor" || n == "forecolor" || n == "hwnd" || n == "tabindex" || n == "setfocus" || n.compare(0, 4, "font") == 0)) { *found = 0; return r; }
        if ((c->kind == K_SHAPE || c->kind == K_LINE) && !(n == "visible" || n == "left" || n == "top" || n == "width" || n == "height" || n == "move" || n == "container" || n == "zorder" || n == "refresh")) { *found = 0; return r; }
        /* a Line has X1 Y1 X2 Y2, not Left Top Width Height Move (VB6: 438) */
        if (c->kind == K_LINE && (n == "left" || n == "top" || n == "width" || n == "height" || n == "move")) { *found = 0; return r; }
    }
    /* the scroll bars have no colours and no font (VB6: 438) */
    if ((c->kind == K_HSCROLL || c->kind == K_VSCROLL) && (n == "backcolor" || n == "forecolor" || n.compare(0, 4, "font") == 0)) {
        *found = 0; return vb_v_empty();
    }
    if (n == "tooltiptext") {
        if (set) { c->tooltip = qs(vb_v2s(val)); if (c->w) c->w->setToolTip(c->tooltip); return vb_v_empty(); }
        return vstr(c->tooltip);
    }
    { vbvar dr; if (dde_member(c, true, n, set, val, &dr)) return dr; }
    if (n == "whatsthishelpid" || n == "helpcontextid") {     /* F1, What's This (help.inc) */
        if (set) { c->misc[n] = vb_v2l(val); return vb_v_empty(); }
        return vb_v_l(ctl_helpid(c, n.c_str()));
    }
    if (n == "showwhatsthis") { whatsthis_ctl(c, QCursor::pos()); return vb_v_empty(); }   /* at the mouse pointer (h1) */
    if (n == "oledragmode" || n == "oledropmode") {
        auto &slot = c->misc[n];
        if (set) { slot = vb_v2l(val); return vb_v_empty(); }
        return vb_v_l(slot);
    }
    /* TabStop: the controls that can have the focus (False: Tab passes them, a click still gives it) */
    if (n == "tabstop" && c->kind != K_LABEL && c->kind != K_IMAGE && c->kind != K_SHAPE && c->kind != K_LINE &&
        c->kind != K_TIMER && c->kind != K_FRAME && c->kind != K_MENU) {
        auto it = c->misc.find(n);
        vb_i32 cur = it != c->misc.end() ? it->second : (c->def ? (propn(c->def->props, "TabStop", 1) ? -1 : 0) : -1);
        if (!set) return vb_v_bool(cur);
        bool on = vb_v2bool(val) != 0;
        c->misc[n] = on ? -1 : 0;
        if (c->w && c->w->focusPolicy() != Qt::NoFocus) c->w->setFocusPolicy(on ? Qt::StrongFocus : Qt::ClickFocus);
        return vb_v_empty();
    }
    /* kept as VB6 has them (what they change is not drawn): Label.UseMnemonic, the mask of graphical buttons */
    if ((n == "usemnemonic" && c->kind == K_LABEL) ||
        ((n == "usemaskcolor" || n == "maskcolor") && (c->kind == K_COMMAND || c->kind == K_CHECK || c->kind == K_OPTION))) {
        static const std::map<std::string, std::pair<const char *, vb_i32>> dflt = {
            { "usemnemonic", { "UseMnemonic", 1 } }, { "usemaskcolor", { "UseMaskColor", 0 } }, { "maskcolor", { "MaskColor", 0xC0C0C0 } } };
        auto d = dflt.at(n);
        auto it = c->misc.find(n);
        vb_i32 cur = it != c->misc.end() ? it->second : (c->def ? (vb_i32)propn(c->def->props, d.first, d.second) : d.second);
        if (n != "maskcolor") cur = cur ? -1 : 0;
        if (!set) return n == "maskcolor" ? vb_v_l(cur) : vb_v_bool(cur);
        c->misc[n] = n == "maskcolor" ? vb_v2l(val) : (vb_v2bool(val) ? -1 : 0);
        return vb_v_empty();
    }
    if (n == "left" || n == "top" || n == "width" || n == "height") {
        double *p = n == "left" ? &c->left : n == "top" ? &c->top : n == "width" ? &c->width : &c->height;
        bool vert = n == "top" || n == "height", pos = n == "left" || n == "top";
        if (set) { *p = to_twips(c, vb_v2d(val), vert, pos); ctl_geometry(c); return vb_v_empty(); }
        return vb_v_f((float)from_twips(c, *p, vert, pos));
    }
    if (n == "visible") {
        if (set) { c->visible = vb_v2bool(val) != 0; if (c->w) c->w->setVisible(c->visible); if (c->align && c->form) align_layout(c->form); if (c->kind == K_USER) uctl_visible(c); return vb_v_empty(); }
        /* VB6 answers whether the control is on the screen: False in Form_Load (t: f_basic) */
        if (c->kind == K_IMAGE || c->kind == K_SHAPE || c->kind == K_LINE) return vb_v_bool(c->visible ? -1 : 0);
        bool on = c->visible && c->form->visible;
        for (vbobj p = c->container; on && p; p = C(p)->container) on = C(p)->visible;
        return vb_v_bool(on ? -1 : 0);
    }
    if (n == "enabled") {
        if (set) { c->enabled = vb_v2bool(val) != 0; if (c->w) c->w->setEnabled(c->enabled); timer_update(c); return vb_v_empty(); }
        return vb_v_bool(c->enabled ? -1 : 0);
    }
    if (n == "tabindex") { if (set) { c->tabindex = vb_v2i(val); return vb_v_empty(); } return vb_v_i((vb_i16)c->tabindex); }
    if (n == "causesvalidation") { if (set) { c->causesval = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(c->causesval ? -1 : 0); }
    /* RightToLeft: as on a Windows that is not BiDi (VB6 on Greek Windows 7, t: f_rtl) - it can be set, it stays False */
    if (n == "righttoleft" && c->kind != K_TIMER && c->kind != K_IMAGE && c->kind != K_SHAPE && c->kind != K_LINE &&
        c->kind != K_MENU) { if (set) return vb_v_empty(); return vb_v_bool(0); }
    if ((c->kind == K_CHECK || c->kind == K_OPTION) && n == "alignment") {   /* 0 left justify, 1 right justify */
        if (set) {
            int a = vb_v2i(val);
            if (a < 0 || a > 1) { vb_raise(380); return vb_v_empty(); }
            if (c->w) { c->w->setProperty("lbalign", a); c->w->setLayoutDirection(a ? Qt::RightToLeft : Qt::LeftToRight); }
            return vb_v_empty();
        }
        return vb_v_i((vb_i16)(c->w ? c->w->property("lbalign").toInt() : 0));
    }
    if (n == "dragmode") { if (set) { c->dragmode = vb_v2i(val); return vb_v_empty(); } return vb_v_i((vb_i16)c->dragmode); }
    if (n == "drag") {
        long act = arg_l(argc, argv, 0, 1);
        if (act == 1) drag_begin(c);
        else if (drag_src == c->self) drag_end(cursor_pos, act == 2);
        return vb_v_empty();
    }
    if (n == "mousepointer") {
        if (!set) return vb_v_i((vb_i16)c->mousepointer);
        c->mousepointer = vb_v2i(val);
        apply_pointer(c->w, c->mousepointer, c->mouseicon);
        return vb_v_empty();
    }
    if (n == "mouseicon") return mouseicon_member(&c->mouseicon, c->w, c->mousepointer, set, val);
    if (n == "container") return vb_vtmp(vb_v_obj(c->container ? c->container : c->form->form));
    if (n == "hwnd" || n == "hdc") return vb_v_l(hwnd_of(c->w));
    if (n == "setfocus") {
        if (!can_focus(c) || !c->form->visible) { vb_raise(5); return vb_v_empty(); }   /* Invalid procedure call or argument */
        focus_to(c->form, c);
        return vb_v_empty();
    }
    if (n == "move") {
        if (argc > 0) c->left = to_twips(c, vb_v2d(argv[0]), false, true);
        if (argc > 1 && !vb_ismissing(argv[1])) c->top = to_twips(c, vb_v2d(argv[1]), true, true);
        if (argc > 2 && !vb_ismissing(argv[2])) c->width = to_twips(c, vb_v2d(argv[2]), false, false);
        if (argc > 3 && !vb_ismissing(argv[3])) c->height = to_twips(c, vb_v2d(argv[3]), true, false);
        ctl_geometry(c);
        return vb_v_empty();
    }
    if (n == "refresh" || n == "zorder") return vb_v_empty();
    if (c->kind != K_TIMER) {
        if (n == "backcolor") {
            if (!set) return vb_v_l((vb_i32)c->back);
            c->back = vb_v2l(val);
            if (c->g) { c->g->back = c->back; gfx_cls(c->g); }
            apply_look(c);
            return vb_v_empty();
        }
        if (n == "forecolor" && c->kind != K_COMMAND) {
            if (!set) return vb_v_l((vb_i32)c->fore);
            c->fore = vb_v2l(val);
            if (c->g) c->g->fore = c->fore;
            apply_look(c);
            return vb_v_empty();
        }
        vbvar r = font_prop(c->self, &c->font, n, set, val, found);
        if (*found || vb_errpend) return r;
        *found = 1;
    }
    switch (c->kind) {
    case K_COMMAND:
        if (n == "caption") { if (set) { c->caption = qs(vb_v2s(val)); if (auto *b = qobject_cast<QPushButton *>(c->w.data())) b->setText(c->caption); return vb_v_empty(); } return vstr(c->caption); }
        if (n == "value") { if (set) { if (vb_v2bool(val)) fire_ctl(c, "click"); return vb_v_empty(); } return vb_v_bool(0); }
        break;
    case K_LABEL: case K_FRAME:
        if (n == "caption") {
            if (set) {
                c->caption = qs(vb_v2s(val));
                if (auto *l = qobject_cast<QLabel *>(c->w.data())) l->setText(label_shown(c));
                if (auto *g = qobject_cast<QGroupBox *>(c->w.data())) g->setTitle(c->caption);
                label_fit(c);
                return vb_v_empty();
            }
            return vstr(c->caption);
        }
        if (c->kind == K_LABEL && (n == "autosize" || n == "wordwrap")) {
            bool &b = n == "autosize" ? c->autosize : c->wordwrap;
            if (!set) return vb_v_bool(b ? -1 : 0);
            b = vb_v2bool(val) != 0;
            label_style(c, qobject_cast<QLabel *>(c->w.data()));
            label_fit(c);
            return vb_v_empty();
        }
        if (c->kind == K_LABEL && n == "alignment") {
            if (!set) return vb_v_i((vb_i16)c->alignment);
            int v = vb_v2i(val);
            if (v < 0 || v > 2) { vb_raise(380); return vb_v_empty(); }
            c->alignment = v;
            label_style(c, qobject_cast<QLabel *>(c->w.data()));
            return vb_v_empty();
        }
        if (c->kind == K_LABEL && n == "backstyle") {
            if (!set) return vb_v_i((vb_i16)c->backstyle);
            c->backstyle = vb_v2i(val) ? 1 : 0;
            apply_look(c);
            return vb_v_empty();
        }
        break;
    case K_TEXT:
        if (n == "text") { if (set) { set_text(c, qs(vb_v2s(val)), true); return vb_v_empty(); } return vstr(get_text(c)); }
        if ((n == "selstart" || n == "sellength" || n == "seltext") && qobject_cast<QPlainTextEdit *>(c->w.data())) {
            /* positions of VB6 count CR LF as two characters; Qt has one (\n) */
            auto *pe = qobject_cast<QPlainTextEdit *>(c->w.data());
            QString q = pe->toPlainText();
            auto vbpos = [&](int qp) { return qp + (int)q.left(qp).count('\n'); };
            auto qpos = [&](int vp) { int i = 0, v = 0; while (i < q.size() && v < vp) { v += q[i] == '\n' ? 2 : 1; i++; } return i; };
            QTextCursor tc = pe->textCursor();
            int a = tc.selectionStart(), b = tc.selectionEnd();
            if (!set) {
                if (n == "selstart") return vb_v_l(vbpos(a));
                if (n == "sellength") return vb_v_l(vbpos(b) - vbpos(a));
                return vstr(q.mid(a, b - a).replace('\n', QStringLiteral("\r\n")));
            }
            if (n == "selstart") { int v = vb_v2l(val); if (v < 0) { vb_raise(380); return vb_v_empty(); } tc.setPosition(qpos(v)); pe->setTextCursor(tc); return vb_v_empty(); }
            if (n == "sellength") {
                int v = vb_v2l(val); if (v < 0) { vb_raise(380); return vb_v_empty(); }
                tc.setPosition(a); tc.setPosition(qpos(vbpos(a) + v), QTextCursor::KeepAnchor); pe->setTextCursor(tc); return vb_v_empty();
            }
            QString ins = qs(vb_v2s(val)).replace(QStringLiteral("\r\n"), QStringLiteral("\n"));
            tc.insertText(ins);                     /* fires Change (textChanged) */
            pe->setTextCursor(tc);
            return vb_v_empty();
        }
        if (n == "selstart" || n == "sellength" || n == "seltext") {
            auto *le = qobject_cast<QLineEdit *>(c->w.data());
            if (!le) break;
            int start = le->hasSelectedText() ? le->selectionStart() : le->cursorPosition();
            int len = le->hasSelectedText() ? (int)le->selectedText().size() : 0;
            if (!set) return n == "selstart" ? vb_v_l(start) : n == "sellength" ? vb_v_l(len) : vstr(le->selectedText());
            if (n == "selstart") { le->setCursorPosition(vb_v2l(val)); return vb_v_empty(); }
            if (n == "sellength") { le->setSelection(start, vb_v2l(val)); return vb_v_empty(); }
            QString t = le->text(); t.replace(start, len, qs(vb_v2s(val)));
            set_text(c, t, true);
            le->setCursorPosition(start + (int)qs(vb_v2s(val)).size());
            return vb_v_empty();
        }
        if (n == "locked") {
            auto *le = qobject_cast<QLineEdit *>(c->w.data());
            auto *pe = qobject_cast<QPlainTextEdit *>(c->w.data());
            if (set) { if (le) le->setReadOnly(vb_v2bool(val) != 0); if (pe) pe->setReadOnly(vb_v2bool(val) != 0); return vb_v_empty(); }
            return vb_v_bool((le && le->isReadOnly()) || (pe && pe->isReadOnly()) ? -1 : 0);
        }
        if (n == "multiline") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_bool(qobject_cast<QPlainTextEdit *>(c->w.data()) ? -1 : 0); }
        if (n == "scrollbars") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_i((vb_i16)propn(c->def->props, "ScrollBars", 0)); }
        if (n == "maxlength") { auto *le = qobject_cast<QLineEdit *>(c->w.data()); if (set) { if (le) le->setMaxLength(vb_v2l(val) ? vb_v2l(val) : 32767); return vb_v_empty(); } return vb_v_l(le && le->maxLength() != 32767 ? le->maxLength() : 0); }
        break;
    case K_CHECK:
        if (n == "caption") { if (set) { c->caption = qs(vb_v2s(val)); if (auto *b = qobject_cast<QCheckBox *>(c->w.data())) b->setText(c->caption); return vb_v_empty(); } return vstr(c->caption); }
        if (n == "value") {
            if (!set) return vb_v_i((vb_i16)c->value);
            int v = vb_v2i(val);
            if (v < 0 || v > 2) { vb_raise(380); return vb_v_empty(); }
            if (v == c->value) return vb_v_empty();
            c->value = v;
            if (auto *b = qobject_cast<QCheckBox *>(c->w.data())) { QSignalBlocker bl(b); b->setTristate(v == 2); b->setCheckState(v == 0 ? Qt::Unchecked : v == 1 ? Qt::Checked : Qt::PartiallyChecked); }
            fire_ctl(c, "click");
            return vb_v_empty();
        }
        break;
    case K_OPTION:
        if (n == "caption") { if (set) { c->caption = qs(vb_v2s(val)); if (auto *b = qobject_cast<QRadioButton *>(c->w.data())) b->setText(c->caption); return vb_v_empty(); } return vstr(c->caption); }
        if (n == "value") {
            auto *rb = qobject_cast<QRadioButton *>(c->w.data());
            if (!set) return vb_v_bool(rb && rb->isChecked() ? -1 : 0);
            bool on = vb_v2bool(val) != 0;
            if (!rb || rb->isChecked() == on) return vb_v_empty();
            { QSignalBlocker bl(rb); if (on) rb->setChecked(true); else { rb->setAutoExclusive(false); rb->setChecked(false); rb->setAutoExclusive(true); } }
            if (on) fire_ctl(c, "click");
            return vb_v_empty();
        }
        break;
    case K_LIST: case K_COMBO: case K_FILE: case K_DRIVE: {
        if (c->kind != K_LIST && c->kind != K_COMBO && (n == "additem" || n == "removeitem" || n == "clear" || n == "text" || n == "itemdata" || n == "newindex")) break;
        vbvar r = list_member(c, n, kind, argc, argv, found);
        if (*found || vb_errpend) return r;
        *found = 1;
        break;
    }
    case K_PICTURE: {
        vbvar r = gfx_member(c->g, n, set, val, argc, argv, found);
        if (*found || vb_errpend) return r;
        *found = 1;
        if (n == "borderstyle") return vb_v_i((vb_i16)propn(c->def->props, "BorderStyle", 1));
        break;
    }
    case K_TIMER:
        if (n == "interval") {
            if (!set) return vb_v_l(c->interval);
            int v = vb_v2l(val);
            if (v < 0 || v > 65535) { vb_raise(380); return vb_v_empty(); }
            c->interval = v; timer_update(c);
            return vb_v_empty();
        }
        break;
    default: break;
    }
    /* what is left of VB6's members: Default / Cancel of a button, the ComboBox's text box, design-time
     * properties read at run time (Appearance, BorderStyle of a Label), DragIcon */
    if ((c->kind == K_COMMAND || (c->uc && propn(fstate(c->uc)->def->props, "DefaultCancel", 0) != 0))
            && (n == "default" || n == "cancel")) {   /* a user control with DefaultCancel: its Extender has them */
        auto it = c->misc.find(n);
        vb_i32 cur = it != c->misc.end() ? it->second : (c->def ? (vb_i32)propn(c->def->props, n == "default" ? "Default" : "Cancel", 0) : 0);
        if (!set) return vb_v_bool(cur ? -1 : 0);
        c->misc[n] = vb_v2bool(val) ? -1 : 0;
        if (n == "default") if (auto *b = qobject_cast<QPushButton *>(c->w.data())) { b->setDefault(vb_v2bool(val) != 0); b->setAutoDefault(vb_v2bool(val) != 0); }
        return vb_v_empty();
    }
    if (c->kind == K_COMBO && (n == "selstart" || n == "sellength" || n == "seltext" || n == "locked")) {
        auto *cb = qobject_cast<QComboBox *>(c->w.data());
        QLineEdit *le = cb ? cb->lineEdit() : nullptr;
        if (n == "locked") {
            auto it = c->misc.find(n);
            if (!set) return vb_v_bool(it != c->misc.end() && it->second ? -1 : 0);
            c->misc[n] = vb_v2bool(val) ? -1 : 0;
            if (le) le->setReadOnly(vb_v2bool(val) != 0);
            if (cb) cb->setEnabled(c->enabled);
            return vb_v_empty();
        }
        if (!le) return n == "seltext" ? vstr(QString()) : vb_v_l(0);
        int start = le->hasSelectedText() ? le->selectionStart() : le->cursorPosition(), len = le->selectedText().size();
        if (!set) return n == "selstart" ? vb_v_l(start) : n == "sellength" ? vb_v_l(len) : vstr(le->selectedText());
        if (n == "selstart") le->setCursorPosition(vb_v2l(val));
        else if (n == "sellength") le->setSelection(start, vb_v2l(val));
        else le->insert(qs(vb_v2s(val)));
        return vb_v_empty();
    }
    if (n == "appearance") { if (set) return vb_v_empty(); return vb_v_i((vb_i16)(c->def ? propn(c->def->props, "Appearance", 1) : 1)); }
    if (n == "borderstyle" && c->kind == K_LABEL) {
        auto it = c->misc.find(n);
        vb_i32 cur = it != c->misc.end() ? it->second : (c->def ? (vb_i32)propn(c->def->props, "BorderStyle", 0) : 0);
        if (!set) return vb_v_i((vb_i16)cur);
        c->misc[n] = vb_v2l(val);
        if (auto *f = qobject_cast<QFrame *>(c->w.data())) f->setFrameStyle(vb_v2l(val) ? (QFrame::Panel | QFrame::Sunken) : QFrame::NoFrame);
        return vb_v_empty();
    }
    if (n == "dragicon") return set ? vb_v_empty() : vb_vtmp(vb_v_obj(new_pic(QImage(), 0)));
    if ((c->kind == K_COMMAND || c->kind == K_CHECK || c->kind == K_OPTION) &&
            (n == "picture" || n == "downpicture" || n == "disabledpicture")) {
        int i = n == "picture" ? 0 : n == "downpicture" ? 1 : 2;
        if (set) {
            vbobj np = val.vt == VT_DISPATCH && val.u.o && pic_of(val.u.o) ? vb_oid(val.u.o) : nullptr;
            if (np) vb_oaddref(np);
            if (c->btnpic[i]) vb_orelease(c->btnpic[i]);
            c->btnpic[i] = np;
            button_icon(c);
            return vb_v_empty();
        }
        return vb_vtmp(vb_v_obj(picture_get(&c->btnpic[i])));
    }
    if (n == "style" && c->def && (c->kind == K_COMMAND || c->kind == K_CHECK || c->kind == K_OPTION ||
                                   c->kind == K_LIST || c->kind == K_COMBO)) {
        if (set) { vb_raise(382); return vb_v_empty(); }     /* read-only at run time */
        return vb_v_i((vb_i16)propn(c->def->props, "Style", 0));
    }
    if (c->kind == K_TEXT && n == "passwordchar") {
        auto it = c->misc.find(n);
        QString cur;
        if (it != c->misc.end()) { if (it->second) cur = QChar((char16_t)it->second); }
        else if (c->def) prop(c->def->props, "PasswordChar", &cur);
        if (!set) return vstr(cur.left(1));
        QString v = qs(vb_v2s(val)).left(1);
        c->misc[n] = v.isEmpty() ? 0 : v[0].unicode();
        if (auto *le = qobject_cast<QLineEdit *>(c->w.data())) le->setEchoMode(v.isEmpty() ? QLineEdit::Normal : QLineEdit::Password);
        return vb_v_empty();
    }
    if (c->kind == K_TEXT && n == "alignment") {
        if (!set) return vb_v_i((vb_i16)c->alignment);
        int v = (int)vb_v2l(val);
        if (v < 0 || v > 2) { vb_raise(380); return vb_v_empty(); }
        c->alignment = v;
        Qt::Alignment a = v == 1 ? Qt::AlignRight : v == 2 ? Qt::AlignHCenter : Qt::AlignLeft;
        if (auto *le = qobject_cast<QLineEdit *>(c->w.data())) le->setAlignment(a | Qt::AlignVCenter);
        return vb_v_empty();
    }
    if (c->kind == K_LIST && n == "columns") {
        auto it = c->misc.find(n);
        vb_i32 cur = it != c->misc.end() ? it->second : (c->def ? (vb_i32)propn(c->def->props, "Columns", 0) : 0);
        if (!set) return vb_v_i((vb_i16)cur);
        vb_i32 v = vb_v2l(val);
        if (vb_errpend) return vb_v_empty();
        if (v < 0 || v > 32767) { vb_raise(380); return vb_v_empty(); }
        c->misc[n] = v;
        list_columns(LW(c), (int)v);
        return vb_v_empty();
    }
    if (c->kind == K_LIST && (n == "selcount" || n == "topindex")) {
        QListWidget *lw = LW(c);
        if (n == "selcount") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_i((vb_i16)(lw ? lw->selectedItems().size() : 0)); }
        if (!set) {
            QListWidgetItem *it = lw ? lw->itemAt(2, 2) : nullptr;
            return vb_v_i((vb_i16)(it ? lw->row(it) : 0));
        }
        int i = (int)vb_v2l(val);
        if (!lw || i < 0 || i >= lw->count()) { if (!(lw && i == 0)) { vb_raise(380); return vb_v_empty(); } return vb_v_empty(); }
        lw->scrollToItem(lw->item(i), QAbstractItemView::PositionAtTop);
        return vb_v_empty();
    }
    if (n == "align" && c->def && c->def->parent < 0 && (c->kind == K_PICTURE || c->kind == K_DATA || c->kind == K_OCX)) {
        if (!set) return vb_v_i((vb_i16)c->align);
        int v = (int)vb_v2l(val);
        if (v < 0 || v > 4) { vb_raise(380); return vb_v_empty(); }
        c->align = v;
        if (c->form) align_layout(c->form);
        return vb_v_empty();
    }
    if (n == "borderstyle") {                    /* a design-time property: read back, kept when set */
        auto it = c->misc.find(n);
        int dflt = (c->kind == K_TEXT || c->kind == K_LIST || c->kind == K_COMBO || c->kind == K_FILE || c->kind == K_DIR ||
                    c->kind == K_DRIVE || c->kind == K_FRAME || c->kind == K_PICTURE) ? 1 : 0;
        vb_i32 cur = it != c->misc.end() ? it->second : (c->def ? (vb_i32)propn(c->def->props, "BorderStyle", dflt) : dflt);
        if (!set) return vb_v_i((vb_i16)cur);
        c->misc[n] = vb_v2l(val);
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- control arrays */
static vbvar arr_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static void arr_clear(vbobj o) { delete (CtlArray *)o->ext; o->ext = nullptr; }
/* For Each oText In Me.txtFields: the elements in the order of their Index (the Data Form Wizard, DEV-9 df3) */
static vb_i32 arr_count(vbobj o) { return (vb_i32)((CtlArray *)vb_oid(o)->ext)->elems.size(); }
static vbvar arr_item(vbobj o, vb_i32 i)
{
    CtlArray *a = (CtlArray *)vb_oid(o)->ext;
    auto it = a->elems.begin();
    std::advance(it, i);
    return vb_vtmp(vb_v_obj(it->second));
}
static const vb_class CLS_ARRAY = { "Object", (vb_i32)sizeof(vbobjh), nullptr, nullptr, arr_clear, nullptr, 0, "item",
                                    arr_count, arr_item, nullptr, arr_member, nullptr };

static vbvar arr_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    CtlArray *a = (CtlArray *)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    (void)kind;
    if (n == "count") return vb_v_i((vb_i16)a->elems.size());
    if (n == "lbound") return vb_v_i((vb_i16)(a->elems.empty() ? 0 : a->elems.begin()->first));
    if (n == "ubound") return vb_v_i((vb_i16)(a->elems.empty() ? -1 : a->elems.rbegin()->first));
    if (n == "item") {
        if (argc < 1 || vb_ismissing(argv[0])) { vb_raise(450); return vb_v_empty(); }
        int i = (int)vb_v2l(argv[0]);           /* (an Empty index is 0, as in VB6) */
        if (vb_errpend) return vb_v_empty();
        auto it = a->elems.find(i);
        if (it == a->elems.end()) { vb_raise_arg(340, std::to_string(i).c_str()); return vb_v_empty(); }
        if (argc > 1) {                         /* Option1(1) = True: the element's default property */
            if (kind == VBM_LET || kind == VBM_SET) return vb_late_o(it->second, "", kind, argc - 1, argv + 1);
            return vb_late_o(it->second, "", VBM_GET, argc - 1, argv + 1);
        }
        return vb_vtmp(vb_v_obj(it->second));
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- collections (Forms, Controls) */
struct Coll { std::vector<vbobj> items; };
static vbvar coll_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    Coll *c = (Coll *)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    (void)kind;
    if (n == "count") return vb_v_i((vb_i16)c->items.size());
    if (n == "item") {
        if (argc > 0 && argv[0].vt == VT_BSTR) {
            QString want = qs(argv[0].u.s).toLower();
            for (vbobj x : c->items) {
                vbvar nm = vb_late_o(x, "name", VBM_GET, 0, nullptr);
                if (qs(vb_v2s(nm)).toLower() == want) return vb_vtmp(vb_v_obj(x));
            }
            vb_raise(730); return vb_v_empty();
        }
        long i = arg_l(argc, argv, 0, -1);
        if (i < 0 || i >= (long)c->items.size()) { vb_raise(9); return vb_v_empty(); }
        return vb_vtmp(vb_v_obj(c->items[(size_t)i]));
    }
    *found = 0;
    return vb_v_empty();
}
static vb_i32 coll_count(vbobj o) { return (vb_i32)((Coll *)vb_oid(o)->ext)->items.size(); }
static vbvar coll_item(vbobj o, vb_i32 i) { return vb_vtmp(vb_v_obj(((Coll *)vb_oid(o)->ext)->items[(size_t)i])); }
static void coll_clear(vbobj o) { Coll *c = (Coll *)o->ext; if (c) { for (vbobj x : c->items) vb_orelease(x); delete c; } o->ext = nullptr; }
static const vb_class CLS_COLL = { "Object", (vb_i32)sizeof(vbobjh), nullptr, nullptr, coll_clear, nullptr, 0, "item",
                                   coll_count, coll_item, nullptr, coll_member, nullptr };
static vbvar ctl_list_item(const std::vector<vbobj> &items, vb_i32 argc, vbvar *argv)
{                                               /* Controls(i), ContainedControls("Name"): an item of the list */
    if (argc > 0 && argv[0].vt == VT_BSTR) {
        QString want = qs(argv[0].u.s).toLower();
        for (vbobj x : items) {
            vbvar nm = vb_late_o(x, "name", VBM_GET, 0, nullptr);
            if (qs(vb_v2s(nm)).toLower() == want) return vb_vtmp(vb_v_obj(x));
        }
        vb_raise(730); return vb_v_empty();
    }
    long i = arg_l(argc, argv, 0, -1);
    if (i < 0 || i >= (long)items.size()) { vb_raise(9); return vb_v_empty(); }
    return vb_vtmp(vb_v_obj(items[(size_t)i]));
}
static vbobj make_coll(const std::vector<vbobj> &items)
{
    vbobj o = vb_onew(&CLS_COLL);
    Coll *c = new Coll;
    for (vbobj x : items) { vb_oaddref(x); c->items.push_back(x); }
    o->ext = c;
    return o;
}

/* ---------------------------------------------------------------- fonts of forms and controls */
static VFont *printer_font();
static VFont *font_of(vbobj owner)
{
    owner = vb_oid(owner);
    if (owner->cls->form) return &fstate(owner)->font;
    if (!strcmp(owner->cls->name, "Printer")) return printer_font();
    Ctl *c = C(owner);
    return c ? &c->font : nullptr;
}
static void font_changed(vbobj owner)
{
    owner = vb_oid(owner);
    if (owner->cls->form) { FormState *st = fstate(owner); if (st->canvas) st->canvas->setFont(qfont(st->font)); return; }
    if (!strcmp(owner->cls->name, "Printer")) return;
    Ctl *c = C(owner);
    if (c) apply_look(c);
}

/* ---------------------------------------------------------------- loading */
static bool white_kind(Kind k) { return k == K_TEXT || k == K_LIST || k == K_COMBO || k == K_DRIVE || k == K_DIR || k == K_FILE; }
static long default_back(Kind k) { return white_kind(k) ? 0x80000005L : 0x8000000FL; }
static long default_fore(Kind k) { return white_kind(k) ? 0x80000008L : 0x80000012L; }

static vbobj new_ctl_object(FormState *st, const vbf_ctldef *d, Kind k, const std::string &type)
{
    vbobj o = vb_onew(ctl_class(k, type));
    vb_oaddref(o);                              /* the form keeps it */
    Ctl *c = new Ctl;
    c->kind = k; c->type = type; c->form = st; c->self = o; c->def = d;
    c->name = QString::fromUtf8(d->name);
    c->index = d->index;
    const char *const *p = d->props;
    c->left = propn(p, "Left", 0); c->top = propn(p, "Top", 0);
    c->width = propn(p, "Width", k == K_TIMER ? 420 : 1215); c->height = propn(p, "Height", k == K_TIMER ? 420 : 495);
    if (!prop(p, "Width", nullptr) && prop(p, "_ExtentX", nullptr)) {      /* an .ocx control: its size in HiMetric */
        c->width = std::round(propn(p, "_ExtentX", 0) * 1440.0 / 2540.0);
        c->height = std::round(propn(p, "_ExtentY", 0) * 1440.0 / 2540.0);
    }
    c->visible = propn(p, "Visible", -1) != 0;
    c->enabled = propn(p, "Enabled", -1) != 0;
    c->checked = propn(p, "Checked", 0) != 0;
    c->tabindex = (int)propn(p, "TabIndex", 0);
    prop(p, "Tag", &c->tag);
    prop(p, "ToolTipText", &c->tooltip);
    prop(p, "Shortcut", &c->shortcut);
    prop(p, "DataField", &c->datafield);
    prop(p, "DataSource", &c->datasource);
    dfmt_from_props(c, p);
    prop(p, "Caption", &c->caption);            /* not in the .frm: "" (VB6 oracle f_data: a Label) */
    c->value = (int)propn(p, "Value", 0);
    c->interval = (int)propn(p, "Interval", 0);
    c->back = propc(p, "BackColor", default_back(k));
    c->fore = propc(p, "ForeColor", default_fore(k));
    c->font = st->font;                         /* the .frm has a Font only when it is not the form's */
    font_from_props(&c->font, p);
    c->stretch = propn(p, "Stretch", 0) != 0;
    c->autosize = propn(p, "AutoSize", 0) != 0;
    c->shape = (int)propn(p, "Shape", 0);
    c->backstyle = (int)propn(p, "BackStyle", k == K_SHAPE ? 0 : 1);
    c->wordwrap = propn(p, "WordWrap", 0) != 0;
    c->alignment = (int)propn(p, "Alignment", 0);
    if (c->def->parent < 0) c->align = (int)propn(p, "Align", 0);    /* on the form itself */
    c->borderstyle = (int)propn(p, "BorderStyle", 1);
    c->borderwidth = (int)propn(p, "BorderWidth", 1);
    c->bordercolor = propc(p, "BorderColor", 0x80000008L);
    c->fillcolor = propc(p, "FillColor", 0);
    c->fillstyle = (int)propn(p, "FillStyle", 1);
    c->x1 = propn(p, "X1", 0); c->y1 = propn(p, "Y1", 0); c->x2 = propn(p, "X2", 0); c->y2 = propn(p, "Y2", 0);
    c->smin = (int)propn(p, "Min", 0); c->smax = (int)propn(p, "Max", 32767);
    c->sval = (int)propn(p, "Value", 0);
    c->small = (int)propn(p, "SmallChange", 1); c->large = (int)propn(p, "LargeChange", 1);
    if (k == K_DIR || k == K_FILE) c->path = QDir::currentPath();
    prop(p, "Pattern", &c->pattern);
    c->f_archive = propn(p, "Archive", -1) != 0; c->f_hidden = propn(p, "Hidden", 0) != 0;
    c->f_normal = propn(p, "Normal", -1) != 0; c->f_readonly = propn(p, "ReadOnly", -1) != 0;
    c->f_system = propn(p, "System", 0) != 0;
    c->causesval = propn(p, "CausesValidation", -1) != 0;
    c->dragmode = (int)propn(p, "DragMode", 0);
    c->mousepointer = (int)propn(p, "MousePointer", 0);
    o->ext = c;
    st->ctls.push_back(o);
    return o;
}

/* Align: the controls at the edges of the client area, in the order of the form; on an MDIForm the
 * children have the rest */
static void align_layout(FormState *st)
{
    if (!st->w || aligning) return;
    bool any = false;
    for (vbobj co : st->ctls) if (C(co)->align && C(co)->w) any = true;
    if (!any) return;
    aligning = true;
    int x0 = 0, y0 = 0, W, H;
    if (st->mdi) {
        QMenuBar *mb = st->w->layout() ? qobject_cast<QMenuBar *>(st->w->layout()->menuBar()) : nullptr;
        y0 = mb && mb->isVisible() ? mb->sizeHint().height() : 0;
        W = st->w->width(); H = st->w->height() - y0;
    } else {
        if (!st->canvas) { aligning = false; return; }
        QRect cr = st->canvas->contentsRect();
        W = cr.width(); H = cr.height();
        /* before the form is shown (Form_Load) its canvas has no size yet: the client area of the .frm
         * (VB6 oracle f_cool: an aligned CoolBar is 7200 wide in Form_Load) */
        if (!st->shown_once) { W = (int)std::lround(st->clientw / TPP); H = (int)std::lround(st->clienth / TPP); }
    }
    int l = 0, r = W, t = 0, b = H;
    /* the controls aligned to the same side are stacked in the order of their places in the .frm (the order VB6
     * keeps: a Toolbar at Top 0 above a PictureBox at Top 540 written before it, aw1 frmBrowser / Explorer); one
     * that gets Align later goes after them */
    std::vector<vbobj> order;
    for (vbobj co : st->ctls) {
        Ctl *c = C(co);
        if (!c->align || !c->w) continue;
        if (!c->align_keyed) {
            c->align_keyed = true;
            c->align_key = c->align == 1 ? c->top : c->align == 2 ? -(c->top + c->height) :
                           c->align == 3 ? c->left : -(c->left + c->width);
            if (st->shown_once) c->align_key = 1e12;
        }
        order.push_back(co);
    }
    for (int side = 1; side <= 4; side++) {     /* sort each side in its own places */
        std::vector<size_t> places;
        std::vector<vbobj> same;
        for (size_t i = 0; i < order.size(); i++) if (C(order[i])->align == side) { places.push_back(i); same.push_back(order[i]); }
        std::stable_sort(same.begin(), same.end(), [](vbobj a, vbobj b) { return C(a)->align_key < C(b)->align_key; });
        for (size_t k = 0; k < places.size(); k++) order[places[k]] = same[k];
    }
    for (vbobj co : order) {
        Ctl *c = C(co);
        if (!c->align || !c->w) continue;
        if (st->mdi && c->w->parentWidget() != st->w) { c->w->setParent(st->w); c->w->setVisible(c->visible); }
        if (!c->visible) continue;
        int w = (int)(c->width / TPP), h = (int)(c->height / TPP);
        QRect g;
        switch (c->align) {
        case 1: g = QRect(l, t, r - l, h); t += h; break;
        case 2: g = QRect(l, b - h, r - l, h); b -= h; break;
        case 3: g = QRect(l, t, w, b - t); l += w; break;
        default: g = QRect(r - w, t, w, b - t); r -= w; break;
        }
        c->left = g.x() * TPP; c->top = g.y() * TPP; c->width = g.width() * TPP; c->height = g.height() * TPP;
        c->w->setGeometry(g.translated(x0, y0));
        if (c->g) gfx_resize(c->g, g.width() - 2 * border_px(c), g.height() - 2 * border_px(c));
        c->w->raise();
        if (c->ocx) ocx_resized(c);
    }
    if (st->mdi && st->w->layout()) st->w->layout()->setContentsMargins(l, t, W - r, H - b);
    aligning = false;
}

static void apply_after_widget(Ctl *c)
{
    const char *const *p = c->def->props;
    QString v;
    if (c->kind == K_TEXT) {
        auto *le = qobject_cast<QLineEdit *>(c->w.data());
        if (le) { QSignalBlocker b(le); le->setText(prop(p, "Text", &v) ? v : QString()); le->setCursorPosition(0); }
        else if (prop(p, "Text", &v)) set_text(c, v, false);
    } else if (c->kind == K_COMBO || c->kind == K_LIST) {
        if (prop(p, "List", &v) && !v.isEmpty()) {
            QStringList items = v.split(QChar(0x1F));
            if (auto *lw = LW(c)) lw->addItems(items);
            if (auto *cb = CB(c)) { QSignalBlocker b(cb); QSignalBlocker b2(cb->lineEdit()); cb->addItems(items); }
            c->itemdata.assign(items.size(), 0);
            QString d;
            if (prop(p, "ItemData", &d) && !d.isEmpty()) {
                QStringList nums = d.split(QChar(0x1F));
                for (int i = 0; i < nums.size() && i < (int)c->itemdata.size(); i++) c->itemdata[i] = nums[i].toLong();
            }
        }
        if (auto *cb = CB(c)) { QSignalBlocker b(cb); QSignalBlocker b2(cb->lineEdit()); cb->setCurrentIndex(-1); cb->setEditText(prop(p, "Text", &v) ? v : QString()); }
    } else if (c->kind == K_IMAGE || c->kind == K_PICTURE) {
        QImage img = frx_image(p, "Picture");
        if (!img.isNull()) set_picture(c, new_pic(img, frx_pictype(p, "Picture")));
    } else if (c->kind == K_COMMAND || c->kind == K_CHECK || c->kind == K_OPTION) {
        static const char *const pn[3] = { "Picture", "DownPicture", "DisabledPicture" };
        for (int i = 0; i < 3; i++) {
            QImage img = frx_image(p, pn[i]);
            if (!img.isNull()) { c->btnpic[i] = new_pic(img, frx_pictype(p, pn[i])); vb_oaddref(c->btnpic[i]); }
        }
        button_icon(c);
    } else if (c->kind == K_DIR) {
        dir_fill(c);
    } else if (c->kind == K_FILE) {
        file_fill(c);
    } else if (c->kind == K_DRIVE) {
        drive_fill(c);
    } else if (c->kind == K_CHECK) {
        if (auto *b = qobject_cast<QCheckBox *>(c->w.data())) { QSignalBlocker bl(b); b->setCheckState(c->value == 1 ? Qt::Checked : c->value == 2 ? Qt::PartiallyChecked : Qt::Unchecked); }
    } else if (c->kind == K_OPTION) {
        if (auto *b = qobject_cast<QRadioButton *>(c->w.data())) { QSignalBlocker bl(b); b->setChecked(propn(p, "Value", 0) != 0); }
    }
    if (c->ocx) ocx_load(c, p);                 /* the .ocx control reads its part of the .frm */
    if (c->kind == K_USER) uctl_after(c);       /* Resize, ReadProperties */
}

static QWidget *container_widget(FormState *st, Ctl *c)
{
    if (c->def->parent >= 0) {
        const vbf_ctldef *pd = &st->def->ctls[c->def->parent];
        auto it = st->byname.find(lower(pd->name));
        if (it != st->byname.end()) {
            vbobj po = it->second;
            if (po->cls == &CLS_ARRAY) { auto *a = (CtlArray *)po->ext; auto e = a->elems.find(pd->index); po = e == a->elems.end() ? nullptr : e->second; }
            if (po) { c->container = po; Ctl *pc = C(po); if (pc->w) return pc->w; }
        }
    }
    return st->canvas;
}

static void set_tab_order(FormState *st)
{
    std::vector<Ctl *> v;
    for (vbobj o : st->ctls) { Ctl *c = C(o); if (c->w) v.push_back(c); }
    std::stable_sort(v.begin(), v.end(), [](Ctl *a, Ctl *b) { return a->tabindex < b->tabindex; });
    for (size_t i = 1; i < v.size(); i++) QWidget::setTabOrder(v[i - 1]->w, v[i]->w);
}


/* ---------------------------------------------------------------- the window frame (Windows 7)
 * Width/Height of a form include the frame that Windows draws; VB6 programs compute with
 * Width - ScaleWidth, so the runtime keeps the sizes of Windows 7 for each BorderStyle,
 * whatever the window manager of Linux draws (f_win). */
struct Frame { int l, t, r, b; };               /* pixels */
static Frame frame_of(FormState *st)
{
    Frame f;
    switch (st->bstyle) {
    case 0: f = { 0, 0, 0, 0 }; break;
    case 1: case 3: f = { 3, 25, 3, 3 }; break;     /* f_win: 90 / 420 twips, Left = ClientLeft - 45 */
    case 4: f = { 3, 21, 3, 3 }; break;             /* f_win: 90 / 360 */
    case 5: f = { 8, 26, 8, 8 }; break;
    default: f = { 8, 30, 8, 8 }; break;
    }
    if (st->menubar && st->bstyle != 0 && !st->mdichild) f.t += 20;
    return f;
}
static double frame_w(FormState *st) { Frame f = frame_of(st); return (f.l + f.r) * TPP; }
static double frame_h(FormState *st) { Frame f = frame_of(st); return (f.t + f.b) * TPP; }

/* the client area changes size: the widgets, the scale, and the Resize event (sent at once, like WM_SIZE) */
static void set_client(vbobj o, double cw, double ch, bool event)
{
    FormState *st = fstate(o);
    if (cw < 0) cw = 0;
    if (ch < 0) ch = 0;
    bool changed = cw != st->clientw || ch != st->clienth;
    st->clientw = cw; st->clienth = ch;
    int mbh = st->menubar && !st->mdichild ? MENU_PX : 0;
    if (st->w) {
        QSignalBlocker b(st->w);
        bool was = st->sizing;
        st->sizing = true;                      /* Windows resizes the canvas at once: no Resize of its own */
        st->w->resize((int)(cw / TPP), (int)(ch / TPP) + mbh);
        st->sizing = was;
        if (st->sub) {                          /* the frame of the subwindow around the client area */
            QStyle *sy = st->sub->style();
            int fwid = sy->pixelMetric(QStyle::PM_MdiSubWindowFrameWidth, nullptr, st->sub);
            int th = sy->pixelMetric(QStyle::PM_TitleBarHeight, nullptr, st->sub);
            st->sub->resize((int)(cw / TPP) + 2 * fwid, (int)(ch / TPP) + th + fwid);
        }
    }
    gfx_resize(&st->g, (int)(cw / TPP), (int)(ch / TPP));
    st->winw = cw + frame_w(st); st->winh = ch + frame_h(st);
    if (changed && event && st->loaded && st->shown_once) fire_form(o, "resize");
}
/* the window gets its real size in whole pixels when it is shown: ClientHeight 2000 -> 1995 (f_win) */
static void snap_client(vbobj o)
{
    FormState *st = fstate(o);
    set_client(o, std::floor(st->clientw / TPP) * TPP, std::floor(st->clienth / TPP) * TPP, false);
    st->norm_cw = st->clientw; st->norm_ch = st->clienth;
}
/* Width/Height/Move: the window is that size in pixels; the client area is what the frame leaves */
static void set_window(vbobj o, double w, double h, bool event)
{
    FormState *st = fstate(o);
    Frame f = frame_of(st);
    double cw = (std::floor(w / TPP) - f.l - f.r) * TPP, ch = (std::floor(h / TPP) - f.t - f.b) * TPP;
    set_client(o, cw, ch, event);
    st->winw = w; st->winh = h;
    st->norm_cw = st->clientw; st->norm_ch = st->clienth;
}

static void place_window(FormState *st)
{
    if (!st->w || st->mdichild) return;
    Frame f = frame_of(st);
    QRect scr = QGuiApplication::primaryScreen() ? QGuiApplication::primaryScreen()->availableGeometry() : QRect(0, 0, 1024, 768);
    double w = st->clientw + frame_w(st), h = st->clienth + frame_h(st);
    switch (st->startpos) {
    case 0:                                     /* manual: ClientLeft/ClientTop of the .frm */
        st->left = propn(st->def->props, "ClientLeft", 60) - f.l * TPP;
        st->top = propn(st->def->props, "ClientTop", 345) - f.t * TPP;
        break;
    case 1: case 2:                             /* center owner / screen */
        st->left = std::floor((scr.width() * TPP - w) / 2 / TPP) * TPP;
        st->top = std::floor((scr.height() * TPP - h) / 2 / TPP) * TPP;
        break;
    default: {                                  /* Windows default: a cascade */
        static int n;
        st->left = (26 * (n % 8) + 26) * TPP; st->top = (26 * (n % 8) + 26) * TPP;
        n++;
        break;
    }
    }
    st->w->move((int)(st->left / TPP) + f.l, (int)(st->top / TPP) + f.t);
}

/* ---------------------------------------------------------------- MDI */
static vbobj *mdi_slot;
static const vb_class *mdi_cls;
extern "C" void vbf_mdi_register(vbobj *slot, const vb_class *cls) { mdi_slot = slot; mdi_cls = cls; }
/* the MDIForm of the project: the instance that is loaded (Set fMainForm = New frmMain, as the Application Wizard
 * writes it, aw25: not a second MDIForm for the children), else the predeclared one */
static vbobj mdi_live;                          /* weak: loaded_forms holds it while it is loaded */
static vbobj mdi_form()
{
    if (mdi_live) return mdi_live;
    return mdi_slot ? vbf_default(mdi_slot, mdi_cls) : nullptr;
}
static vbobj mdi_current() { return mdi_live ? mdi_live : mdi_slot ? *mdi_slot : nullptr; }
static void set_window(vbobj o, double w, double h, bool event);
static QSize mdi_client_size(FormState *ps)
{
    if (ps->area && ps->area->isVisible()) return ps->area->viewport()->size();
    int h = (int)(ps->clienth / TPP), w = (int)(ps->clientw / TPP);
    for (vbobj co : ps->ctls) {                 /* the aligned controls take their part */
        Ctl *c = C(co);
        if (c->align == 1 || c->align == 2) h -= (int)(c->height / TPP);
        else if (c->align == 3 || c->align == 4) w -= (int)(c->width / TPP);
    }
    return QSize(std::max(w, 50), std::max(h, 50));
}
/* Windows places a new child on a cascade and makes a sizable one as big as the cascade leaves
 * (the user32 rule: spacing = caption + frame, a third of the height in steps) */
static void mdi_child_default(vbobj o)
{
    FormState *st = fstate(o);
    vbobj p = mdi_current();
    if (!p) return;
    QSize a = mdi_client_size(fstate(p));
    const int sp = 25;
    int nst = std::max(1, a.height() / (3 * sp));
    static int n;
    int k = n++ % (nst + 1);
    st->left = sp * k * TPP; st->top = sp * k * TPP;
    if (st->bstyle == 2) set_window(o, (a.width() - sp * nst) * TPP, (a.height() - sp * nst) * TPP, false);
    if (st->sub) st->sub->move(sp * k, sp * k);
}
static void load_form(vbobj o);

/* the active MDI child changes: ActiveForm at once; Deactivate of the old one and Activate of
 * the new one come from the message queue, after the MDIForm's own Activate (f_win) */
/* the menus of the active child replace those of the MDIForm (when it has a visible one) */
static void mdi_menus(vbobj parent)
{
    FormState *ps = fstate(parent);
    if (!ps->w || !ps->w->layout()) return;
    QMenuBar *want = ps->menubar;
    if (vbobj ch = ps->active_child) {
        FormState *cs = fstate(ch);
        if (cs->loaded && cs->visible && cs->menubar)
            for (QAction *a : cs->menubar->actions()) if (a->isVisible()) { want = cs->menubar; break; }
    }
    for (vbobj f : loaded_forms) {
        FormState *fs = fstate(f);
        if (fs->menubar && fs->mdichild && fs->menubar != want) fs->menubar->hide();
    }
    if (ps->menubar && ps->menubar != want) ps->menubar->hide();
    QLayout *lay = ps->w->layout();
    if (lay->menuBar() != want) lay->setMenuBar(want);
    if (want) want->show();
    align_layout(ps);
}
static void child_activate(vbobj parent, vbobj child)
{
    FormState *ps = fstate(parent);
    if (ps->active_child == child) return;
    vbobj old = ps->active_child;
    ps->active_child = child;
    mdi_menus(parent);
    FormState *cs = child ? fstate(child) : nullptr;
    if (cs && cs->sub && ps->area) { QSignalBlocker b(ps->area); ps->area->setActiveSubWindow(cs->sub); }
    if (old) vb_oaddref(old);
    if (child) vb_oaddref(child);
    post([old, child]() {
        if (old) { if (fstate(old)->loaded) fire_form(old, "deactivate"); vb_orelease(old); }
        if (child) {
            FormState *c = fstate(child);
            if (c->loaded && c->visible) {
                Ctl *first = c->focus ? nullptr : first_focusable(c);
                if (first) focus_to(c, first, false);
                fire_form(child, "activate");
                if (first && first->form && c->focus == first->self) fire_ctl(first, "gotfocus");
            }
            vb_orelease(child);
        }
    });
}
static void child_gone(vbobj child)             /* hidden or unloaded: another child becomes active */
{
    vbobj p = mdi_current();
    if (!p) return;
    FormState *ps = fstate(p);
    if (ps->active_child != child) return;
    ps->active_child = nullptr;
    mdi_menus(p);
    vbobj next = nullptr;
    if (ps->area)
        for (QMdiSubWindow *sw : ps->area->subWindowList(QMdiArea::ActivationHistoryOrder))
            if (auto *fw = dynamic_cast<FormWidget *>(sw->widget()))
                if (fw->form != child && std::find(loaded_forms.begin(), loaded_forms.end(), fw->form) != loaded_forms.end() &&
                    fstate(fw->form)->loaded && fstate(fw->form)->visible) next = fw->form;
    if (next) child_activate(p, next);
}

/* the controls of a form or a user control (containers come before what they contain in the .frm) */
static void build_ctls(FormState *st)
{
    const vbf_formdef *d = st->def;
    for (int i = 0; i < d->nctls; i++) {
        const vbf_ctldef *cd = &d->ctls[i];
        std::string type; Kind k = kind_of(cd->cls, &type);
        vbobj co = new_ctl_object(st, cd, k, type);
        Ctl *c = C(co);
        std::string key = lower(cd->name);
        if (cd->index >= 0) {
            auto it = st->byname.find(key);
            vbobj ao;
            if (it == st->byname.end()) {
                ao = vb_onew(&CLS_ARRAY); vb_oaddref(ao);
                CtlArray *a = new CtlArray; a->name = c->name; a->form = st; ao->ext = a;
                st->byname[key] = ao;
            } else ao = it->second;
            ((CtlArray *)ao->ext)->elems[cd->index] = co;
        } else st->byname[key] = co;
        make_widget(c, container_widget(st, c));
        apply_after_widget(c);
    }
    /* the .frm lists the controls from the top of the z-order down; the graphical ones (Label,
     * Shape, Line, Image) are under the others, as Windows draws them on their container */
    auto light = [](Ctl *c) { return c->kind == K_LABEL || c->kind == K_SHAPE || c->kind == K_LINE || c->kind == K_IMAGE; };
    for (int pass = 0; pass < 2; pass++)
        for (auto it = st->ctls.rbegin(); it != st->ctls.rend(); ++it) {
            Ctl *c = C(*it);
            if (c->w && light(c) == (pass == 0) && c->kind != K_MENU) c->w->raise();
        }
}

static void load_form(vbobj o)
{
    o = vb_oid(o);
    FormState *st = fstate(o);
    if (st->loaded) return;
    if (st->def && st->def->mdi == 2) { uctl_load(o); return; }   /* a user control: inside its form */
    ensure_app();
    const vbf_formdef *d = st->def;
    st->mdi = d->mdi != 0;
    if (d->mdi == 1 && !mdi_live) mdi_live = o;            /* this instance is the MDIForm of the children */
    st->mdichild = !st->mdi && propn(d->props, "MDIChild", 0) != 0;
    if (st->mdichild) {                         /* an MDI child loads its MDIForm first */
        vbobj p = mdi_form();
        if (!p) { vb_raise(366); return; }      /* No MDI form available to load */
        load_form(p);
        if (vb_errpend || !fstate(p)->loaded) return;
    }
    st->loaded = true; st->visible = false; st->shown_once = false; st->focus = nullptr;
    st->ctls.clear(); st->byname.clear();
    vb_oaddref(o);
    loaded_forms.push_back(o);
    FormWidget *fw = new FormWidget(o);
    st->w = fw;
    st->bstyle = st->mdi ? 2 : (int)propn(d->props, "BorderStyle", 2);
    st->controlbox = propn(d->props, "ControlBox", -1) != 0;
    st->minbtn = propn(d->props, "MinButton", st->bstyle == 2 ? -1 : 0) != 0;
    st->maxbtn = propn(d->props, "MaxButton", st->bstyle == 2 ? -1 : 0) != 0;
    st->taskbar = propn(d->props, "ShowInTaskbar", st->bstyle == 0 || st->bstyle >= 3 ? 0 : -1) != 0;
    st->moveable = propn(d->props, "Moveable", -1) != 0;
    st->startpos = st->mdichild ? 3 : (int)propn(d->props, "StartUpPosition", 3);
    st->wstate = (int)propn(d->props, "WindowState", 0);
    st->autoshow = propn(d->props, "AutoShowChildren", -1) != 0;
    st->whatsthis = propn(d->props, "WhatsThisHelp", 0) != 0;
    st->whatsthisbtn = propn(d->props, "WhatsThisButton", 0) != 0;
    st->active_child = nullptr;
    if (!st->mdichild) {
        Qt::WindowFlags fl = Qt::Window;
        switch (st->bstyle) {
        case 0: fl |= Qt::FramelessWindowHint; break;
        case 3: fl = Qt::Dialog | Qt::WindowTitleHint | Qt::WindowCloseButtonHint; break;
        case 4: case 5: fl = Qt::Tool | Qt::WindowTitleHint | Qt::WindowCloseButtonHint; break;
        default:
            fl |= Qt::WindowTitleHint | Qt::WindowSystemMenuHint;
            if (st->controlbox) fl |= Qt::WindowCloseButtonHint;
            if (st->minbtn) fl |= Qt::WindowMinimizeButtonHint;
            if (st->maxbtn) fl |= Qt::WindowMaximizeButtonHint;
            break;
        }
        /* the ? of the title bar: WhatsThisHelp and WhatsThisButton, a fixed dialog or a window without
         * Minimize and Maximize (VB6); where the window manager draws it */
        if (st->whatsthis && st->whatsthisbtn && st->controlbox &&
            (st->bstyle == 3 || ((st->bstyle == 1 || st->bstyle == 2) && !st->minbtn && !st->maxbtn)))
            fl |= Qt::WindowContextHelpButtonHint;
        fw->setWindowFlags(fl);
    }
    if (!prop(d->props, "Caption", &st->caption)) st->caption = QString::fromUtf8(d->name);
    fw->setWindowTitle(st->caption);
    {                                           /* Icon of the .frm */
        QImage ic = frx_image(d->props, "Icon");
        if (!ic.isNull()) {
            if (st->icon) vb_orelease(st->icon);
            st->icon = new_pic(ic, 3); vb_oaddref(st->icon);
            fw->setWindowIcon(QIcon(QPixmap::fromImage(ic)));
        }
    }
    st->clientw = propn(d->props, "ClientWidth", 4800); st->clienth = propn(d->props, "ClientHeight", 3600);
    st->left = propn(d->props, "ClientLeft", 60); st->top = propn(d->props, "ClientTop", 345);
    st->keypreview = propn(d->props, "KeyPreview", 0) != 0;
    st->helpctx = (vb_i32)propn(d->props, "HelpContextID", 0);
    st->font = VFont();
    font_from_props(&st->font, d->props);
    st->g = Gfx();
    st->g.back = propc(d->props, "BackColor", 0x8000000F);
    st->g.fore = propc(d->props, "ForeColor", 0x80000012);
    st->g.font = &st->font;
    gfx_from_props(&st->g, d->props, st->clientw, st->clienth);
    st->g.pic = frx_image(d->props, "Picture");
    if (st->picture) { vb_orelease(st->picture); st->picture = nullptr; }
    if (!st->g.pic.isNull()) { st->picture = new_pic(st->g.pic, 1); vb_oaddref(st->picture); gfx_clear(&st->g); }
    st->mousepointer = (int)propn(d->props, "MousePointer", 0);
    {
        QImage mi = frx_image(d->props, "MouseIcon");
        if (!mi.isNull() && !st->mouseicon) { st->mouseicon = new_pic(mi, 3); vb_oaddref(st->mouseicon); }
    }
    if (st->mousepointer) apply_pointer(fw, st->mousepointer, st->mouseicon);
    auto *lay = new QVBoxLayout(fw);
    lay->setContentsMargins(0, 0, 0, 0);
    lay->setSpacing(0);
    Canvas *cv = new Canvas(&st->g, fw);
    cv->setAttribute(Qt::WA_StaticContents);    /* a smaller window has nothing new to paint: no Paint (VB6, t21 geo) */
    cv->setFont(qfont(st->font));
    cv->setFocusPolicy(Qt::StrongFocus);
    lay->addWidget(cv);
    st->canvas = cv;
    if (st->mdi) {                              /* the client area of an MDIForm holds the children */
        cv->hide();
        auto *area = new QMdiArea(fw);
        area->setBackground(qcolor(propc(d->props, "BackColor", 0x8000000C)));
        lay->addWidget(area);
        st->area = area;
        QObject::connect(area, &QMdiArea::subWindowActivated, [o](QMdiSubWindow *sw) {
            if (!sw) return;
            if (auto *cfw = dynamic_cast<FormWidget *>(sw->widget())) child_activate(o, cfw->form);
        });
    }
    cv->on_paint = [o]() { if (fstate(o)->activate_pending) activate_now(o); fire_form(o, "paint"); };
    cv->on_resize = [o](int w, int h) {
        FormState *s = fstate(o);
        if (s->sizing) { align_layout(s); return; }   /* set_client / WindowState give the size and the Resize */
        double cw = w * TPP, ch = h * TPP;
        bool changed = cw != s->clientw || ch != s->clienth;
        s->clientw = cw; s->clienth = ch;
        align_layout(s);
        if (changed && s->loaded && s->visible && s->shown_once) fire_form(o, "resize");
    };
    build_ctls(st);
    build_menus(st);
    if (st->mdichild && st->menubar) {          /* shown in the MDIForm while the child is active */
        st->w->layout()->setMenuBar(nullptr);
        st->menubar->hide();
        st->menubar->setParent(fstate(mdi_form())->w);
    }
    if (st->mdichild) {
        FormState *ps = fstate(mdi_form());
        QSignalBlocker blk(ps->area);
        st->sub = ps->area->addSubWindow(fw);
        st->sub->setWindowTitle(st->caption);
        st->sub->hide();
    }
    set_client(o, st->clientw, st->clienth, false);
    snap_client(o);                             /* the window has its size in pixels at Load already, and the scale
                                                 * too (VB6: ScaleHeight 1995 of ClientHeight 2000 in Form_Load, t21 geo;
                                                 * Height - ScaleHeight 420 of a fixed form, f_win) */
    st->norm_cw = st->clientw; st->norm_ch = st->clienth;
    if (st->mdichild) mdi_child_default(o);
    place_window(st);
    align_layout(st);
    set_tab_order(st);
    for (vbobj co : st->ctls) timer_update(C(co));
    ocx_loaded(st);                             /* a Toolbar finds the ImageList written after it (aw24) */
    if (!vb_errpend) adodc_after_load(st);      /* an ADO Data control opens before Load (VB6 oracle f_dfmt3) */
    fire_form(o, "load");
    if (st->loaded && !vb_errpend) data_after_load(st);    /* the Data controls open their recordsets */
    if (st->mdichild && st->loaded && !vb_errpend && fstate(mdi_form())->autoshow) form_show(o, 0);
}

/* Windows activates another window of the program when the active one is hidden or
 * unloaded: the form that was active before it (f_basic: Form1 Activate after Form2.Hide) */
static std::vector<vbobj> activation_order;     /* most recent last (weak) */
/* the Activate of a form that is shown: after the code that showed it (posted); a paint that comes
 * first (Windows sends WM_PAINT of a new window before the posted message: f_win in Windows 7) has it
 * done first, as Windows activates a window before it paints it */
static void activate_now(vbobj o)
{
    FormState *s = fstate(o);
    s->activate_pending = false;
    if (s->loaded && s->visible && active_form != o) {
        vbobj prev = active_form;
        active_form = o;
        activation_order.erase(std::remove(activation_order.begin(), activation_order.end(), o), activation_order.end());
        activation_order.push_back(o);
        if (s->w) s->w->activateWindow();
        /* the first control in the tab order has the focus already in Form_Activate,
         * its GotFocus comes after it (f_ui) */
        Ctl *first = s->focus ? nullptr : first_focusable(s);
        if (first) focus_to(s, first, false);
        if (prev && fstate(prev)->loaded) fire_form(prev, "deactivate");
        fire_form(o, "activate");
        if (first && first->form && s->focus == first->self) fire_ctl(first, "gotfocus");
    }
}
static void schedule_activate(vbobj o)
{
    fstate(o)->activate_pending = true;
    vb_oaddref(o);
    post([o]() {
        if (fstate(o)->activate_pending) activate_now(o);
        vb_orelease(o);
    });
}
static void lost_active(vbobj o, bool owned_unload = false)
{
    activation_order.erase(std::remove(activation_order.begin(), activation_order.end(), o), activation_order.end());
    if (active_form != o) return;
    active_form = nullptr;
    for (auto it = activation_order.rbegin(); it != activation_order.rend(); ++it) {
        FormState *s = fstate(*it);
        if (s->loaded && s->visible) {
            schedule_activate(*it);
            /* Windows: when a window without a taskbar button (ShowInTaskbar = False) is
             * destroyed, the next one is activated, deactivated and activated again (f_win) */
            if (owned_unload) {
                vbobj n = *it;
                vb_oaddref(n);
                post([n]() {
                    if (active_form == n && fstate(n)->loaded) { fire_form(n, "deactivate"); fire_form(n, "activate"); }
                    vb_orelease(n);
                });
            }
            return;
        }
    }
}

/* Load, Show or a member of a form that is not loaded: a form unloaded by its own Form_Load is
 * "Object was unloaded" to the code that loaded it (VB6 oracle do2: Load Form1 whose Load ends with
 * Unload Me); the startup form is shown by vbf_start, without it */
static bool load_checked(vbobj o)
{
    o = vb_oid(o);
    FormState *st = fstate(o);
    if (st->loaded) return true;
    if (st->def && st->def->mdi == 2) { load_form(o); return !vb_errpend; }
    load_form(o);
    if (vb_errpend) return false;
    if (!st->loaded) { vb_raise(364); return false; }
    return true;
}

static void form_show(vbobj o, int modal)
{
    o = vb_oid(o);
    FormState *st = fstate(o);
    load_form(o);
    if (vb_errpend || !st->loaded) return;
    if (!st->visible) {
        st->visible = true;
        if (st->mdichild) {
            vbobj p = mdi_form();
            if (p && !fstate(p)->visible) form_show(p, 0);
            {
                vbobj pf = mdi_form();
                QSignalBlocker b(fstate(pf)->area);
                if (st->w) st->w->show();
                if (st->sub) st->sub->show();
            }
            if (!st->shown_once) { st->shown_once = true; snap_client(o); fire_form(o, "resize"); }
            ocx_shown(st);
            child_activate(mdi_form(), o);
        } else {
            if (st->w) {
                QSignalBlocker blk(st->area.data());   /* an MDIForm: its children activate the VB6 way */
                if (st->wstate == 2) st->w->showMaximized();
                else if (st->wstate == 1) st->w->showMinimized();
                else st->w->show();
            }
            if (!st->shown_once) { st->shown_once = true; snap_client(o); fire_form(o, "resize"); }
            uctl_shown(st);                     /* the Show of its user controls: before Activate */
            ocx_shown(st);                      /* a CoolBar: its bands take their children */
            /* Activate comes from the message loop, after the code that showed the form */
            schedule_activate(o);
        }
    }
    if (modal) {
        extern std::vector<vbobj> modal_stack;
        modal_stack.push_back(o);
        struct Pop { ~Pop() { if (!modal_stack.empty()) modal_stack.pop_back(); } } pop;
        QEventLoop loop;
        QTimer poll;
        QObject::connect(&poll, &QTimer::timeout, [&]() { FormState *s = fstate(o); if (!s->loaded || !s->visible) loop.quit(); });
        poll.start(20);
        if (monkey_on) monkey_start();          /* the robot: frmX.Show vbModal in Sub Main never reaches vbf_run */
        loop.exec();
    }
}

static void form_hide(vbobj o)
{
    o = vb_oid(o);
    FormState *st = fstate(o);
    load_form(o);
    if (!st->loaded) return;
    st->visible = false;
    if (st->sub) { st->sub->hide(); child_gone(o); return; }
    if (st->w) st->w->hide();
    lost_active(o);
}

static bool unload_query(vbobj o, int mode)
{
    vbvar a[2] = { vb_v_i(0), vb_v_i((vb_i16)mode) };
    vb_i32 m = vb_mark();
    fire(o, ev_prefix(o) + "queryunload", 2, a);
    int cancel = vb_v2i(a[0]);
    vb_release(m);
    return !cancel && !vb_errpend;
}

static void unload_destroy(vbobj o)            /* the Unload event; then the form goes */
{
    FormState *st = fstate(o);
    vbvar b[1] = { vb_v_i(0) };
    vb_i32 m = vb_mark();
    fire(o, ev_prefix(o) + "unload", 1, b);
    int cancel = vb_v2i(b[0]);
    vb_release(m);
    if (cancel || vb_errpend) return;
    st->unloading = true;
    if (st->mdichild) child_gone(o); else lost_active(o, !st->taskbar && active_form == o);
    st->focus = nullptr;
    for (vbobj co : st->ctls) { Ctl *c = C(co); if (c && c->kind == K_DATA) data_unload(c); }
    for (vbobj co : st->ctls) { Ctl *c = C(co); if (c) { if (c->timer) c->timer->stop(); c->form = nullptr; } }
    if (st->sub) {
        QMdiSubWindow *sw = st->sub;
        st->sub = nullptr;
        if (st->w) { st->w->hide(); st->w->setParent(nullptr); st->w->deleteLater(); st->w = nullptr; }
        sw->hide(); sw->deleteLater();
    }
    if (st->canvas) { st->canvas->on_paint = nullptr; st->canvas->on_resize = nullptr; }   /* no events for a form that went */
    if (st->w) { st->w->hide(); st->w->deleteLater(); st->w = nullptr; }
    st->menubar = nullptr;
    std::vector<vbobj> ctls; ctls.swap(st->ctls);
    std::map<std::string, vbobj> names; names.swap(st->byname);
    st->loaded = false; st->visible = false; st->unloading = false; st->active_child = nullptr;
    if (mdi_live == o) mdi_live = nullptr;
    for (auto &kv : names) if (kv.second->cls == &CLS_ARRAY) vb_orelease(kv.second);
    for (vbobj co : ctls) vb_orelease(co);
    auto it = std::find(loaded_forms.begin(), loaded_forms.end(), o);
    if (it != loaded_forms.end()) { loaded_forms.erase(it); vb_orelease(o); }   /* may run Form_Terminate */
    if (loaded_forms.empty() && running && app) QTimer::singleShot(0, []() { if (loaded_forms.empty()) app->quit(); });
}

static void unload_form(vbobj o, int mode)
{
    o = vb_oid(o);
    FormState *st = fstate(o);
    if (!st->loaded || st->unloading || st->in_unload) return;
    if (!st->mdi) {
        vb_oaddref(o);                          /* the form lives on until its own Unload code is done */
        st->in_unload = true;
        if (unload_query(o, mode)) unload_destroy(o);
        fstate(o)->in_unload = false;
        vb_orelease(o);
        return;
    }
    st->in_unload = true;
    /* an MDIForm: its QueryUnload, the children's QueryUnload (vbFormMDIForm), their Unload, its Unload */
    vb_oaddref(o);
    std::vector<vbobj> kids;
    for (vbobj f : loaded_forms) if (fstate(f)->mdichild) { kids.push_back(f); vb_oaddref(f); }
    bool ok = unload_query(o, mode);
    for (vbobj k : kids) if (ok && fstate(k)->loaded) ok = unload_query(k, 4);
    for (vbobj k : kids) if (ok && fstate(k)->loaded) { unload_destroy(k); if (fstate(k)->loaded) ok = false; }
    for (vbobj &k : kids) { vb_orelease(k); k = nullptr; }
    if (ok) unload_destroy(o);
    fstate(o)->in_unload = false;
    vb_orelease(o);
}



/* ---------------------------------------------------------------- the Controls collection of a form (Add, Remove) */
static vbvar ctls_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vbobj ctls_form(vbobj o) { return (vbobj)vb_oid(o)->ext; }
static vb_i32 ctls_count(vbobj o) { FormState *st = fstate(ctls_form(o)); return (vb_i32)st->ctls.size(); }
static vbvar ctls_item(vbobj o, vb_i32 i) { FormState *st = fstate(ctls_form(o)); return vb_vtmp(vb_v_obj(st->ctls[(size_t)i])); }
static void ctls_clear(vbobj o) { if (o->ext) vb_orelease((vbobj)o->ext); o->ext = nullptr; }
static const vb_class CLS_CTLS = { "Object", (vb_i32)sizeof(vbobjh), nullptr, nullptr, ctls_clear, nullptr, 0, "item",
                                   ctls_count, ctls_item, nullptr, ctls_member, nullptr, "_newenum" };
/* For Each c In Form.Controls goes over the controls the form has when the loop starts: a loop that adds controls
 * (a shadow Shape for each control) does not go over the new ones (robot 6.8: GiveShadow looped for ever) */
struct CtlSnap { std::vector<vbobj> v; };
static vb_i32 snap_count(vbobj o) { auto *s = (CtlSnap *)vb_oid(o)->ext; return s ? (vb_i32)s->v.size() : 0; }
static vbvar snap_item(vbobj o, vb_i32 i)
{
    auto *s = (CtlSnap *)vb_oid(o)->ext;
    if (!s || i < 0 || i >= (vb_i32)s->v.size()) return vb_v_empty();
    return vb_vtmp(vb_v_obj(s->v[(size_t)i]));
}
static void snap_clear(vbobj o)
{
    auto *s = (CtlSnap *)o->ext;
    o->ext = nullptr;
    if (s) { for (vbobj c : s->v) vb_orelease(c); delete s; }
}
static const vb_class CLS_CTLSNAP = { "Object", (vb_i32)sizeof(vbobjh), nullptr, nullptr, snap_clear, nullptr, 0, nullptr,
                                      snap_count, snap_item, nullptr, nullptr, nullptr };
static vbobj controls_of(vbobj form)
{
    vbobj o = vb_onew(&CLS_CTLS);
    vb_oaddref(form);
    o->ext = form;
    return o;
}
static vbvar ctls_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    if (!strcmp(name, "_newenum")) {             /* For Each: the controls of now */
        *found = 1;
        FormState *sst = fstate(ctls_form(o));
        vbobj sn = vb_onew(&CLS_CTLSNAP);
        auto *cs = new CtlSnap;
        for (vbobj c : sst->ctls) { vb_oaddref(c); cs->v.push_back(c); }
        sn->ext = cs;
        return vb_vtmp(vb_v_obj(sn));
    }
    vbobj form = ctls_form(o);
    FormState *st = fstate(form);
    std::string n = lower(name);
    (void)kind;
    *found = 1;
    if (n == "count") return vb_v_i((vb_i16)st->ctls.size());
    if (n == "item") {
        if (argc > 0 && argv[0].vt == VT_BSTR) {
            auto it = st->byname.find(lower(qs(argv[0].u.s).toUtf8().constData()));
            if (it == st->byname.end()) { vb_raise(730); return vb_v_empty(); }
            return vb_vtmp(vb_v_obj(it->second));
        }
        long i = arg_l(argc, argv, 0, -1);
        if (i < 0 || i >= (long)st->ctls.size()) { vb_raise(9); return vb_v_empty(); }
        return vb_vtmp(vb_v_obj(st->ctls[(size_t)i]));
    }
    if (n == "add") {                           /* Controls.Add("VB.TextBox", "txtNew" [, container]) */
        QString progid = arg_s(argc, argv, 0), cname = arg_s(argc, argv, 1);
        std::string key = lower(cname.toUtf8().constData());
        if (cname.isEmpty()) { vb_raise(5); return vb_v_empty(); }
        if (st->byname.count(key)) { vb_raise(727); return vb_v_empty(); }   /* There is already a control with the name */
        std::string type; Kind k = kind_of(progid.toUtf8().constData(), &type);
        if (k == K_OTHER || k == K_MENU) { vb_raise(711); return vb_v_empty(); }   /* Invalid class string */
        static const char *const no_props[] = { nullptr };
        auto *d = new vbf_ctldef;
        d->cls = strdup(progid.toUtf8().constData()); d->name = strdup(cname.toUtf8().constData());
        d->index = -1; d->parent = -1; d->props = no_props;
        vbobj co = new_ctl_object(st, d, k, type);
        Ctl *c = C(co);
        c->added = true;
        c->visible = false;                     /* a new control is invisible, 1200 x 480, no caption (f_misc) */
        c->width = 1200; c->height = 480;
        if (k == K_USER) {                      /* a user control: the size of its designer (f_uctl2) */
            const vb_class *uc = uctl_class(progid.toUtf8().constData());
            const vbf_formdef *ud = uc ? (const vbf_formdef *)uc->form : nullptr;
            if (ud) { c->width = propn(ud->props, "ClientWidth", 4800); c->height = propn(ud->props, "ClientHeight", 3600); }
        }
        c->caption.clear();
        int maxtab = -1;
        for (vbobj x : st->ctls) if (x != co) maxtab = std::max(maxtab, C(x)->tabindex);
        c->tabindex = maxtab + 1;
        st->byname[key] = co;
        QWidget *parent = st->canvas;
        if (argc > 2 && argv[2].vt == VT_DISPATCH && argv[2].u.o) {
            vbobj po = vb_oid(argv[2].u.o);
            if (!po->cls->form && C(po) && C(po)->w) { c->container = po; parent = C(po)->w; }
        }
        make_widget(c, parent);
        apply_after_widget(c);
        timer_update(c);
        return vb_vtmp(vb_v_obj(co));
    }
    if (n == "remove") {
        vbobj co = nullptr;
        if (argc > 0 && argv[0].vt == VT_DISPATCH) co = vb_oid(argv[0].u.o);
        else if (argc > 0) { auto it = st->byname.find(lower(arg_s(argc, argv, 0).toUtf8().constData())); if (it != st->byname.end()) co = it->second; }
        if (!co) { vb_raise(730); return vb_v_empty(); }      /* Control not found */
        Ctl *c = C(co);
        if (!c || !c->added) { vb_raise(729); return vb_v_empty(); }   /* f_misc: a control of the .frm */
        st->byname.erase(lower(c->name.toUtf8().constData()));
        if (st->focus == co) st->focus = nullptr;
        if (c->timer) c->timer->stop();
        if (c->w) { c->w->hide(); c->w->deleteLater(); c->w = nullptr; }
        c->form = nullptr;
        auto ci = std::find(st->ctls.begin(), st->ctls.end(), co);
        if (ci != st->ctls.end()) { st->ctls.erase(ci); vb_orelease(co); }
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- hWnd: numbers for the windows of forms and controls */
static std::map<vb_i32, QPointer<QWidget>> hwnds;
static int hwnd_of(QWidget *w)
{
    if (!w) return 0;
    QVariant v = w->property("vbhwnd");
    if (v.isValid()) return v.toInt();
    static vb_i32 next = 0x00010010;
    vb_i32 h;
#ifdef _WIN32
    if (w->isWindow()) h = (vb_i32)(intptr_t)w->winId();    /* a form: its real HWND, for the Declares of user32 */
    else
#endif
    { h = next; next += 4; }
    w->setProperty("vbhwnd", h);
    hwnds[h] = w;
    return h;
}
static QWidget *widget_of_hwnd(vb_i32 h)
{
    auto it = hwnds.find(h);
    return it == hwnds.end() ? nullptr : it->second.data();
}

/* ---------------------------------------------------------------- the form's built-in members */
extern "C" vbvar vbf_form_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    o = vb_oid(o);
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    FormState *st = fstate(o);
    if (st->def && st->def->mdi == 2) {         /* UserControl.Extender, .Ambient ... */
        vbvar r = uctl_self_member(o, n, set, val, argc, argv, found);
        if (*found || vb_errpend) return r;
        *found = 1;
    }
    if (n == "show") { if (load_checked(o)) form_show(o, (int)arg_l(argc, argv, 0, 0)); return vb_v_empty(); }
    if (n == "hide") { form_hide(o); return vb_v_empty(); }
    if (!load_checked(o)) return vb_v_empty(); /* any other member loads the form */
    if (n == "name") return vstr(QString::fromUtf8(st->def->name));
    if (n == "caption") {
        if (set) { st->caption = qs(vb_v2s(val)); if (st->w) st->w->setWindowTitle(st->caption); if (st->sub) st->sub->setWindowTitle(st->caption); return vb_v_empty(); }
        return vstr(st->caption);
    }
    if (n == "width" || n == "height") {
        bool wd = n == "width";
        if (!set) {                             /* before it is shown, Height has no menu bar yet (VB6: 2565, then
                                                 * 2865 at the Resize of Show, t21 geo in Windows 7) */
            double v = wd ? st->winw : st->winh;
            if (!wd && !st->shown_once && st->menubar && !st->mdichild && st->bstyle != 0) v -= MENU_PX * TPP;
            return vb_v_f((float)v);
        }
        double v = vb_v2d(val);
        if (st->wstate != 0) { vb_raise(384); return vb_v_empty(); }   /* can't be changed while minimized or maximized */
        if (wd) set_window(o, v, st->winh, true); else set_window(o, st->winw, v, true);
        return vb_v_empty();
    }
    if (n == "left" || n == "top") {
        double *p = n == "left" ? &st->left : &st->top;
        if (!set) return vb_v_f((float)*p);
        *p = vb_v2d(val);
        if (st->w && !st->mdichild) { Frame f = frame_of(st); st->w->move((int)(st->left / TPP) + f.l, (int)(st->top / TPP) + f.t); }
        else if (st->sub) st->sub->move((int)(st->left / TPP), (int)(st->top / TPP));
        return vb_v_empty();
    }
    if (n == "move") {
        if (st->wstate != 0) { vb_raise(384); return vb_v_empty(); }
        if (given(argc, argv, 0)) st->left = vb_v2d(argv[0]);
        if (given(argc, argv, 1)) st->top = vb_v2d(argv[1]);
        if (st->w && !st->mdichild) { Frame f = frame_of(st); st->w->move((int)(st->left / TPP) + f.l, (int)(st->top / TPP) + f.t); }
        else if (st->sub) st->sub->move((int)(st->left / TPP), (int)(st->top / TPP));
        set_window(o, given(argc, argv, 2) ? vb_v2d(argv[2]) : st->winw, given(argc, argv, 3) ? vb_v2d(argv[3]) : st->winh, true);
        return vb_v_empty();
    }
    if (n == "windowstate") {
        if (!set) return vb_v_i((vb_i16)st->wstate);
        int v = vb_v2i(val);
        if (v < 0 || v > 2) { vb_raise(380); return vb_v_empty(); }
        if (v == st->wstate) return vb_v_empty();
        if (st->wstate == 0) { st->norm_cw = st->clientw; st->norm_ch = st->clienth; }
        st->wstate = v;
        if (v == 1) {                           /* minimized: Resize with ScaleWidth = ScaleHeight = 0 (f_win) */
            if (st->w && st->visible) { if (st->sub) st->sub->showMinimized(); else st->w->showMinimized(); }
            st->clientw = st->clienth = 0;
            st->g.sw = st->g.sh = 0;
            if (st->visible) fire_form(o, "resize");
        } else if (v == 2) {
            QRect scr = QGuiApplication::primaryScreen() ? QGuiApplication::primaryScreen()->availableGeometry() : QRect(0, 0, 1024, 768);
            if (st->w && st->visible) {         /* Windows resizes it at once: one Resize, with the new sizes */
                st->sizing = true;
                if (st->sub) st->sub->showMaximized(); else st->w->showMaximized();
                st->sizing = false;
            }
            Frame f = frame_of(st);
            set_client(o, scr.width() * TPP, (scr.height() - (f.t - f.b)) * TPP, st->visible);   /* the frame goes outside the screen */
        } else {
            if (st->w && st->visible) {
                st->sizing = true;
                if (st->sub) st->sub->showNormal(); else st->w->showNormal();
                st->sizing = false;
            }
            set_client(o, st->norm_cw, st->norm_ch, st->visible);   /* Width/Height: the real size again (4995) */
        }
        return vb_v_empty();
    }
    if (n == "borderstyle") { if (set) return vb_v_empty(); return vb_v_i((vb_i16)st->bstyle); }   /* f_win: no error */
    if (n == "controlbox" || n == "minbutton" || n == "maxbutton" || n == "showintaskbar" || n == "moveable") {
        if (set) { vb_raise(382); return vb_v_empty(); }   /* Set not supported at run time */
        bool v = n == "controlbox" ? st->controlbox : n == "minbutton" ? st->minbtn : n == "maxbutton" ? st->maxbtn : n == "showintaskbar" ? st->taskbar : st->moveable;
        return vb_v_bool(v ? -1 : 0);
    }
    if (n == "mdichild") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_bool(st->mdichild ? -1 : 0); }
    if (n == "autoshowchildren" && st->mdi) { if (set) { st->autoshow = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(st->autoshow ? -1 : 0); }
    if (n == "activeform" && st->mdi) return vb_vtmp(vb_v_obj(st->active_child));
    if (n == "arrange" && st->mdi) {
        if (st->area) { long how = arg_l(argc, argv, 0); if (how == 0) st->area->cascadeSubWindows(); else st->area->tileSubWindows(); }
        return vb_v_empty();
    }
    if (n == "visible") {
        if (set) { if (vb_v2bool(val)) form_show(o, 0); else form_hide(o); return vb_v_empty(); }
        return vb_v_bool(st->visible ? -1 : 0);
    }
    if (n == "enabled") { if (set) { st->enabled = vb_v2bool(val) != 0; if (st->w) st->w->setEnabled(st->enabled); return vb_v_empty(); } return vb_v_bool(st->enabled ? -1 : 0); }
    if (n == "tag") { if (set) { st->tag = qs(vb_v2s(val)); return vb_v_empty(); } return vstr(st->tag); }
    if (n == "printform") return vb_v_empty();  /* the printout of the form: no printer at hand here */
    { vbvar dr; if (dde_member(st, false, n, set, val, &dr)) return dr; }
    if (n == "helpcontextid") { if (set) { st->helpctx = vb_v2l(val); return vb_v_empty(); } return vb_v_l(st->helpctx); }
    if (n == "whatsthishelp") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_bool(st->whatsthis ? -1 : 0); }   /* read-only at run time (f_help: 382) */
    if (n == "whatsthisbutton") { if (set) { vb_raise(382); return vb_v_empty(); } return vb_v_bool(st->whatsthisbtn ? -1 : 0); }
    if (n == "whatsthismode") { help_whatsthis_mode(); return vb_v_empty(); }
    if (n == "oledropmode") return set ? vb_v_empty() : vb_v_l(0);
    if (n == "keypreview") { if (set) { st->keypreview = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(st->keypreview ? -1 : 0); }
    if (n == "appearance" && !st->host) {      /* read-only at run time (VB6: 382) */
        if (set) { vb_raise(382); return vb_v_empty(); }
        return vb_v_i((vb_i16)propn(st->def->props, "Appearance", 1));
    }
    if (n == "righttoleft") { if (set) return vb_v_empty(); return vb_v_bool(0); }   /* (not BiDi: stays False) */
    if (n == "backcolor") {
        if (!set) return vb_v_l((vb_i32)st->g.back);
        st->g.back = vb_v2l(val);
        gfx_cls(&st->g);
        return vb_v_empty();
    }
    if (n == "forecolor") { if (!set) return vb_v_l((vb_i32)st->g.fore); st->g.fore = vb_v2l(val); return vb_v_empty(); }
    if (n == "picture") {
        if (!set) return vb_vtmp(vb_v_obj(picture_get(&st->picture)));
        vbobj np = val.vt == VT_DISPATCH && pic_of(val.u.o) ? vb_oid(val.u.o) : new_pic(QImage(), 0);
        vb_oaddref(np);
        if (st->picture) vb_orelease(st->picture);
        st->picture = np;
        st->g.pic = pic_of(np)->img;
        gfx_clear(&st->g);
        return vb_v_empty();
    }
    if (n == "activecontrol") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        if (!st->focus) { vb_raise(91); return vb_v_empty(); }
        return vb_vtmp(vb_v_obj(st->focus));
    }
    if (n == "hwnd" || n == "hdc") return vb_v_l(hwnd_of(st->w));     /* hDC: a handle for the Declares */
    if (n == "count" && !set) return vb_v_i((vb_i16)st->ctls.size());   /* Form.Count: of its controls */
    if (n == "controls" || (n.empty() && argc > 0 && !set)) {   /* Controls("Text1"), Controls(0), Form!Text1: an item */
        vbobj cs = controls_of(o);
        if (argc > 0 && !set) return vb_late_o(cs, "item", VBM_GET, argc, argv);
        return vb_vtmp(vb_v_obj(cs));
    }
    if (n == "mousepointer") {
        if (!set) return vb_v_i((vb_i16)st->mousepointer);
        st->mousepointer = vb_v2i(val);
        apply_pointer(st->w, st->mousepointer, st->mouseicon);
        return vb_v_empty();
    }
    if (n == "mouseicon") return mouseicon_member(&st->mouseicon, st->w, st->mousepointer, set, val);
    if (n == "icon") {
        if (!set) return vb_vtmp(vb_v_obj(picture_get(&st->icon)));
        vbobj np = val.vt == VT_DISPATCH && pic_of(val.u.o) ? vb_oid(val.u.o) : new_pic(QImage(), 0);
        vb_oaddref(np);
        if (st->icon) vb_orelease(st->icon);
        st->icon = np;
        if (st->w) st->w->setWindowIcon(pic_of(np)->type ? QIcon(QPixmap::fromImage(pic_of(np)->img)) : QIcon());
        return vb_v_empty();
    }
    if (n == "validatecontrols") {              /* the Validate event of the active control; Cancel: error 380 */
        if (st->focus && !validate(C(st->focus)) && !vb_errpend) vb_raise(380);
        return vb_v_empty();
    }
    if (n == "popupmenu") {
        vbobj mo = argc > 0 && argv[0].vt == VT_DISPATCH ? vb_oid(argv[0].u.o) : nullptr;
        Ctl *mc = mo && mo->cls == &CLS_MENU ? C(mo) : nullptr;
        if (!mc || !mc->menu) { vb_raise(490); return vb_v_empty(); }   /* Object must be a menu with a submenu */
        QPoint at = cursor_pos;
        if (given(argc, argv, 2) && given(argc, argv, 3) && st->canvas)
            at = st->canvas->mapToGlobal(QPoint(ipx(px_x(&st->g, vb_v2d(argv[2]))), ipx(px_y(&st->g, vb_v2d(argv[3])))));
        flush_pending();
        mc->menu->exec(at);                     /* modal: the Click of the chosen item runs inside */
        return vb_v_empty();
    }
    if (n == "refresh") {                       /* AutoRedraw = False: the Paint event at once */
        if (!st->g.autoredraw && st->canvas && !painting) { painting = true; fire_form(o, "paint"); painting = false; }
        if (st->canvas) st->canvas->update();
        return vb_v_empty();
    }
    if (n == "zorder") { if (st->w) { if (arg_l(argc, argv, 0)) st->w->lower(); else st->w->raise(); } return vb_v_empty(); }
    if (n == "setfocus") {
        if (!st->visible) { vb_raise(5); return vb_v_empty(); }
        if (st->mdichild) { vbobj p = mdi_form(); if (p) child_activate(p, o); return vb_v_empty(); }
        if (st->w) st->w->activateWindow();
        if (active_form != o) schedule_activate(o);
        return vb_v_empty();
    }
    {
        vbvar r = gfx_member(&st->g, n, set, val, argc, argv, found);
        if (*found || vb_errpend) return r;
        r = font_prop(o, &st->font, n, set, val, found);
        if (*found || vb_errpend) return r;
        *found = 1;
    }
    auto it = st->byname.find(n);              /* a control by name: Form2.Label1 */
    if (it != st->byname.end()) {
        if (argc && !set) return vb_late_o(it->second, "", VBM_GET, argc, argv);
        if (set) return vb_late_o(it->second, "", kind, argc, argv);
        return vb_vtmp(vb_v_obj(it->second));
    }
    *found = 0;
    return vb_v_empty();
}

extern "C" void vbf_form_clear(vbobj o)
{
    FormState *st = (FormState *)o->ext;
    if (!st) return;
    if (st->w) st->w->deleteLater();
    if (st->picture) vb_orelease(st->picture);
    if (st->icon) vb_orelease(st->icon);
    delete st;
    o->ext = nullptr;
}

/* ---------------------------------------------------------------- C API */
extern "C" vbobj vbf_default(vbobj *slot, const vb_class *cls)
{
    return vb_odefault(slot, cls);              /* in the slot before Form_Initialize: Form1.Height there is this form */
}

extern "C" vbobj vbf_ctl(vbobj form, const char *name)
{
    form = vb_oid(form);
    if (!form) { vb_raise(91); return nullptr; }
    load_form(form);
    FormState *st = fstate(form);
    auto it = st->byname.find(name);
    if (it == st->byname.end()) { vb_raise(91); return nullptr; }
    return it->second;
}

extern "C" void vbf_load(vbobj o)
{
    o = vb_oid(o);
    if (!o) { vb_raise(91); return; }
    if (o->cls->form) load_checked(o);
}

extern "C" void vbf_load_elem(vbobj arr, vb_i32 index)
{
    arr = vb_oid(arr);
    if (!arr || arr->cls != &CLS_ARRAY) { vb_raise(343); return; }    /* Object not an array */
    CtlArray *a = (CtlArray *)arr->ext;
    if (a->elems.count(index)) { vb_raise(360); return; }              /* Object already loaded */
    if (a->elems.empty() || !a->form) { vb_raise(91); return; }
    Ctl *src = C(a->elems.begin()->second);
    FormState *st = a->form;
    vbobj bag = src->kind == K_USER ? uctl_copy_bag(src) : nullptr;
    vbobj co = new_ctl_object(st, src->def, src->kind, src->type);
    Ctl *c = C(co);
    c->bag = bag;
    c->index = index; c->visible = false;      /* a new element is invisible, like in VB6 */
    c->left = src->left; c->top = src->top; c->width = src->width; c->height = src->height;
    c->caption = src->caption; c->enabled = src->enabled;
    c->back = src->back; c->fore = src->fore; c->font = src->font;
    a->elems[index] = co;
    make_widget(c, container_widget(st, c));
    apply_after_widget(c);
    if (c->kind == K_OPTION) if (auto *b = qobject_cast<QRadioButton *>(c->w.data())) { QSignalBlocker bl(b); b->setChecked(false); }
    timer_update(c);
}

extern "C" void vbf_unload(vbobj o)
{
    o = vb_oid(o);
    if (!o) { vb_raise(91); return; }
    if (o->cls->form) { unload_form(o, 1); return; }                  /* vbFormCode */
    Ctl *c = C(o);
    if (!c || !c->form || c->index < 0) { vb_raise(362); return; }    /* Can't unload controls created at design time */
    FormState *st = c->form;
    auto it = st->byname.find(lower(c->name.toUtf8().constData()));
    if (it == st->byname.end()) return;
    CtlArray *a = (CtlArray *)it->second->ext;
    for (int i = 0; i < st->def->nctls; i++)   /* elements from the .frm cannot be unloaded */
        if (!strcasecmp(st->def->ctls[i].name, c->def->name) && st->def->ctls[i].index == c->index) { vb_raise(362); return; }
    a->elems.erase(c->index);
    if (st->focus == o) st->focus = nullptr;
    if (c->timer) c->timer->stop();
    if (c->uc) uctl_release(c);                 /* Hide (if it was shown), Terminate */
    /* hidden and out of the form at once: the form shown before the deletion would size it, its picture gone
     * (robot: Extract Icons unloads an element in Form_Load) */
    if (c->w) { c->w->hide(); c->w->setParent(nullptr); c->w->deleteLater(); c->w = nullptr; }
    c->form = nullptr;
    auto ci = std::find(st->ctls.begin(), st->ctls.end(), o);
    if (ci != st->ctls.end()) { st->ctls.erase(ci); vb_orelease(o); }
}


/* ---------------------------------------------------------------- Screen */
static int screen_pointer;
static vbvar screen_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)o;
    ensure_app();
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    QScreen *scr = QGuiApplication::primaryScreen();
    QRect g = scr ? scr->geometry() : QRect(0, 0, 1024, 768);
    if (n == "twipsperpixelx" || n == "twipsperpixely") return vb_v_f((float)TPP);
    if (n == "width") return vb_v_f((float)(g.width() * TPP));
    if (n == "height") return vb_v_f((float)(g.height() * TPP));
    if (n == "activeform") {                    /* an MDIForm: its active child (f_win) */
        vbobj f = active_form;
        if (f && fstate(f)->mdi && fstate(f)->active_child) f = fstate(f)->active_child;
        return vb_vtmp(vb_v_obj(f));
    }
    if (n == "activecontrol") {                 /* of the active form: of the active child of an MDIForm */
        vbobj f = active_form;
        if (f && fstate(f)->mdi && fstate(f)->active_child) f = fstate(f)->active_child;
        if (!f || !fstate(f)->focus) { vb_raise(91); return vb_v_empty(); }
        return vb_vtmp(vb_v_obj(fstate(f)->focus));
    }
    if (n == "mouseicon") {                     /* Screen.MouseIcon: the cursor of MousePointer 99 */
        static vbobj screen_icon = nullptr;
        return mouseicon_member(&screen_icon, nullptr, screen_pointer, set, val);
    }
    if (n == "mousepointer") {
        if (!set) return vb_v_i((vb_i16)screen_pointer);
        int v = vb_v2i(val);
        if (screen_pointer) QApplication::restoreOverrideCursor();
        screen_pointer = v;
        static const Qt::CursorShape shapes[] = { Qt::ArrowCursor, Qt::ArrowCursor, Qt::CrossCursor, Qt::IBeamCursor, Qt::ArrowCursor,
            Qt::SizeAllCursor, Qt::SizeBDiagCursor, Qt::SizeVerCursor, Qt::SizeFDiagCursor, Qt::SizeHorCursor, Qt::UpArrowCursor,
            Qt::WaitCursor, Qt::ForbiddenCursor, Qt::BusyCursor, Qt::WhatsThisCursor, Qt::SizeAllCursor };
        if (v > 0 && v < 16) QApplication::setOverrideCursor(QCursor(shapes[v]));
        return vb_v_empty();
    }
    if (n == "fontcount") return vb_v_i((vb_i16)font_families().size());
    if (n == "fonts") { QStringList f = font_families(); long i = arg_l(argc, argv, 0); return vstr(i >= 0 && i < f.size() ? f[(int)i] : QString()); }
    *found = 0;
    return vb_v_empty();
}
static const vb_class CLS_SCREEN = { "Screen", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr,
                                     nullptr, nullptr, nullptr, screen_member, nullptr };
extern "C" vbobj vbf_screen(void)
{
    static vbobj scr;
    if (!scr) { scr = vb_onew(&CLS_SCREEN); vb_oaddref(scr); }
    return scr;
}

/* ---------------------------------------------------------------- Clipboard */
static vbvar clip_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)o; (void)kind;
    ensure_app();
    std::string n = lower(name);
    QClipboard *cb = QGuiApplication::clipboard();
    *found = 1;
    const QMimeData *md = cb->mimeData();
    long fmt = arg_l(argc, argv, n == "settext" || n == "setdata" ? 1 : 0, n == "setdata" || n == "getdata" ? 2 : 1);
    if (n == "clear") { cb->clear(); return vb_v_empty(); }
    if (n == "settext") {
        auto *m = new QMimeData;
        if (md && md->hasImage()) m->setImageData(md->imageData());
        if (fmt == -16639) m->setData("text/rtf", arg_s(argc, argv, 0).toUtf8()); else m->setText(arg_s(argc, argv, 0));
        cb->setMimeData(m);
        return vb_v_empty();
    }
    if (n == "gettext") {
        if (fmt == -16639) return vstr(md && md->hasFormat("text/rtf") ? QString::fromUtf8(md->data("text/rtf")) : QString());
        return vstr(md && md->hasText() ? md->text() : QString());
    }
    if (n == "getformat") {
        bool r = false;
        if (md) switch (fmt) {
            case 1: r = md->hasText(); break;
            case 2: case 8: case 3: r = md->hasImage(); break;
            case -16639: r = md->hasFormat("text/rtf"); break;
            default: r = false; break;
        }
        return vb_v_bool(r ? -1 : 0);
    }
    if (n == "setdata") {
        Pic *p = argc > 0 && argv[0].vt == VT_DISPATCH ? pic_of(argv[0].u.o) : nullptr;
        if (!p) { vb_raise(13); return vb_v_empty(); }
        auto *m = new QMimeData;
        if (md && md->hasText()) m->setText(md->text());
        m->setImageData(p->img);
        cb->setMimeData(m);
        return vb_v_empty();
    }
    if (n == "getdata") {
        QImage img = md && md->hasImage() ? qvariant_cast<QImage>(md->imageData()) : QImage();
        return vb_vtmp(vb_v_obj(new_pic(img.isNull() ? QImage() : img.convertToFormat(QImage::Format_RGB32), 1)));
    }
    *found = 0;
    return vb_v_empty();
}
static const vb_class CLS_CLIP = { "Clipboard", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr,
                                   nullptr, nullptr, nullptr, clip_member, nullptr };
extern "C" vbobj vbf_clipboard(void)
{
    static vbobj c;
    if (!c) { c = vb_onew(&CLS_CLIP); vb_oaddref(c); }
    return c;
}

/* ---------------------------------------------------------------- Licenses: the licenses of the controls a program
 * adds with Controls.Add; the controls of LinBASIC Studio need none, a key is kept as given */
static std::vector<std::pair<QString, QString>> g_licenses;
static vbvar lic_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)o; (void)kind;
    std::string n = lower(name);
    *found = 1;
    auto find = [](vbvar k) -> int {
        vbvar x = vb_v_letval(k);
        if (x.vt == VT_BSTR) { QString s = qs(x.u.s); for (size_t i = 0; i < g_licenses.size(); i++) if (!s.compare(g_licenses[i].first, Qt::CaseInsensitive)) return (int)i; return -1; }
        int i = vb_v2l(x) - 1;
        return i >= 0 && i < (int)g_licenses.size() ? i : -1;
    };
    if (n == "count") return vb_v_l((vb_i32)g_licenses.size());
    if (n == "add") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        QString id = qs(vb_v2s(argv[0])), key = argc > 1 && !vb_ismissing(argv[1]) ? qs(vb_v2s(argv[1])) : QString();
        int i = find(argv[0]);
        if (i >= 0) { vb_raise(732); return vb_v_empty(); }
        g_licenses.push_back({ id, key });
        return vstr(key);
    }
    if (n == "remove") {
        int i = argc ? find(argv[0]) : -1;
        if (i < 0) { vb_raise(731); return vb_v_empty(); }
        g_licenses.erase(g_licenses.begin() + i);
        return vb_v_empty();
    }
    if (n == "clear") { g_licenses.clear(); return vb_v_empty(); }
    if (n == "item" || n.empty()) {
        int i = argc ? find(argv[0]) : -1;
        if (i < 0) { vb_raise(9); return vb_v_empty(); }
        return vstr(g_licenses[i].second);
    }
    *found = 0;
    return vb_v_empty();
}
static const vb_class CLS_LICENSES = { "Licenses", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item",
                                       nullptr, nullptr, nullptr, lic_member, nullptr };
extern "C" vbobj vbf_licenses(void)
{
    static vbobj c;
    if (!c) { c = vb_onew(&CLS_LICENSES); vb_oaddref(c); }
    return c;
}


/* ---------------------------------------------------------------- Printer: pages drawn like a PictureBox, EndDoc writes a PDF
 * The file: $VBRT_PRINTER_FILE, or <App.Path>/<EXEName>-printer.pdf (A4, 192 dpi pages). */
struct PrinterState {
    std::vector<QImage> pages;
    Gfx g;
    VFont font;
    int orientation = 1, copies = 1, page = 1;
    bool started = false;
    QString lastfile;
};
static PrinterState *prn;
static void prn_page_setup()
{
    Gfx &g = prn->g;
    g.tpp = 7.5;
    int w = 1588, h = 2245;                      /* A4 at 192 dpi */
    if (prn->orientation == 2) std::swap(w, h);
    g.back = 0xFFFFFF; g.fore = 0;
    g.img = QImage(); g.pic = QImage();
    gfx_resize(&g, w, h);
    g.img.fill(Qt::white);
    g.cx = g.cy = 0;
}
static QString prn_file()
{
    const char *e = getenv("VBRT_PRINTER_FILE");
    if (e && *e) return QString::fromUtf8(e);
    vb_i32 m = vb_mark();
    QString r = qs(vb_app_path()) + "/" + qs(vb_app_exename()) + "-printer.pdf";
    vb_release(m);
    return r;
}
static void prn_enddoc(bool write)
{
    if (!prn->started) return;
    prn->pages.push_back(prn->g.img);
    if (write) {
        QString f = prn_file();
        QPdfWriter pdf(f);
        pdf.setResolution(192);
        pdf.setPageSize(QPageSize(QPageSize::A4));
        pdf.setPageOrientation(prn->orientation == 2 ? QPageLayout::Landscape : QPageLayout::Portrait);
        pdf.setPageMargins(QMarginsF(0, 0, 0, 0));
        QPainter p(&pdf);
        for (size_t i = 0; i < prn->pages.size(); i++) {
            if (i) pdf.newPage();
            p.drawImage(QRect(0, 0, pdf.width(), pdf.height()), prn->pages[i]);
        }
        p.end();
        prn->lastfile = f;
    }
    prn->pages.clear();
    prn->started = false;
    prn->page = 1;
    prn_page_setup();
}
static vbvar prn_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)o;
    ensure_app();
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (n == "devicename") return vstr("LinBASIC PDF");
    if (n == "drivername") return vstr("PDF");
    if (n == "port") return vstr(prn_file());
    if (n == "enddoc") { prn_enddoc(true); return vb_v_empty(); }
    if (n == "killdoc") { prn_enddoc(false); return vb_v_empty(); }
    if (n == "newpage") {
        prn->started = true;
        prn->pages.push_back(prn->g.img);
        int mode = prn->g.scalemode; double sl = prn->g.sl, st = prn->g.st, sw = prn->g.sw, sh = prn->g.sh;
        prn_page_setup();
        prn->g.scalemode = mode;
        if (mode == 0) { prn->g.sl = sl; prn->g.st = st; prn->g.sw = sw; prn->g.sh = sh; } else gfx_rescale(&prn->g);
        prn->page++;
        return vb_v_empty();
    }
    if (n == "page") return vb_v_i((vb_i16)prn->page);
    if (n == "orientation") {
        if (!set) return vb_v_i((vb_i16)prn->orientation);
        int v = vb_v2i(val);
        if (v != 1 && v != 2) { vb_raise(380); return vb_v_empty(); }
        if (v != prn->orientation) { prn->orientation = v; if (!prn->started) prn_page_setup(); }
        return vb_v_empty();
    }
    if (n == "copies") { if (set) { prn->copies = vb_v2i(val); return vb_v_empty(); } return vb_v_i((vb_i16)prn->copies); }
    if (n == "papersize") return vb_v_i(9);    /* vbPRPSA4 */
    if (n == "duplex" || n == "colormode" || n == "paperbin" || n == "trackdefault" || n == "zoom") {   /* kept, not used */
        static std::map<std::string, vb_i32> kept = { { "duplex", 1 }, { "colormode", 2 }, { "paperbin", 15 }, { "trackdefault", -1 }, { "zoom", 0 } };
        if (set) { kept[n] = vb_v2l(val); return vb_v_empty(); }
        return n == "trackdefault" ? vb_v_bool(kept[n] ? -1 : 0) : vb_v_i((vb_i16)kept[n]);
    }
    if (n == "printquality") return set ? vb_v_empty() : vb_v_i(-4);   /* vbPRPQHigh */
    if (n == "fontcount") return vb_v_i((vb_i16)font_families().size());
    if (n == "fonts") { int i = argc ? vb_v2i(argv[0]) : 0; QStringList f = font_families(); if (i < 0 || i >= f.size()) { vb_raise(381); return vb_v_empty(); } return vstr(f[i]); }
    if (n == "width") return vb_v_l((vb_i32)std::lround(prn->g.pw * prn->g.tpp));
    if (n == "height") return vb_v_l((vb_i32)std::lround(prn->g.ph * prn->g.tpp));
    if (n == "twipsperpixelx" || n == "twipsperpixely") return vb_v_f((float)prn->g.tpp);
    if (n == "colormode") return vb_v_i(2);
    if (n == "forecolor") { if (set) { prn->g.fore = vb_v2l(val); return vb_v_empty(); } return vb_v_l((vb_i32)prn->g.fore); }
    if (n[0] == '$' || n == "cls") prn->started = true;
    vbvar r = gfx_member(&prn->g, n, set, val, argc, argv, found);
    if (*found || vb_errpend) return r;
    return font_prop(o, &prn->font, n, set, val, found);
}
static VFont *printer_font() { return prn ? &prn->font : nullptr; }
static const vb_class CLS_PRINTER = { "Printer", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr,
                                      nullptr, nullptr, nullptr, prn_member, nullptr };
extern "C" vbobj vbf_printer(void)
{
    static vbobj p;
    if (!p) {
        ensure_app();
        prn = new PrinterState;
        prn->font.name = "Arial"; prn->font.size = real_size(10);
        prn->g.font = &prn->font;
        prn_page_setup();
        p = vb_onew(&CLS_PRINTER); vb_oaddref(p);
    }
    return p;
}
extern "C" vbobj vbf_printers(void)
{
    std::vector<vbobj> v { vbf_printer() };
    return make_coll(v);
}

/* ---------------------------------------------------------------- user32: the cursor and the mouse
 * SetCursorPos, GetCursorPos, mouse_event, ClientToScreen ... work on the windows of the
 * program: mouse_event gives the events to the widget under the cursor, as Windows does
 * (capture while a button is down, a double click from two clicks in time) (f_mouse). */
static QPointer<QWidget> mouse_capture, last_press_w;
static Qt::MouseButtons mouse_btns;
static qint64 last_press_ms = -100000;
static QPoint last_press_pos;
static Qt::MouseButton last_press_b;
static quint64 mouse_ts = 1000;

static QWidget *widget_at(QPoint gp)
{
    std::vector<QWidget *> wins;
    if (QWidget *m = QApplication::activeModalWidget()) wins.push_back(m);
    if (active_form && fstate(active_form)->w) wins.push_back(fstate(active_form)->w);
    for (auto it = activation_order.rbegin(); it != activation_order.rend(); ++it) if (fstate(*it)->w) wins.push_back(fstate(*it)->w);
    for (vbobj f : loaded_forms) if (fstate(f)->w && !fstate(f)->mdichild) wins.push_back(fstate(f)->w);
    for (QWidget *w : wins) {
        if (!w->isVisible() || !w->isWindow()) continue;
        QPoint lp = w->mapFromGlobal(gp);
        if (!w->rect().contains(lp)) continue;
        QWidget *c = w->childAt(lp);
        return c ? c : w;
    }
    return nullptr;
}
static void send_mouse(QEvent::Type type, Qt::MouseButton b)
{
    QWidget *target = ((type == QEvent::MouseButtonRelease || (type == QEvent::MouseMove && mouse_btns)) && mouse_capture) ? mouse_capture.data() : widget_at(cursor_pos);
    if (!target) return;
    QEvent::Type t = type;
    if (type == QEvent::MouseButtonPress) {
        qint64 now = QDateTime::currentMSecsSinceEpoch();
        if (last_press_w == target && last_press_b == b && now - last_press_ms <= QApplication::doubleClickInterval()
            && (cursor_pos - last_press_pos).manhattanLength() <= 4) {
            t = QEvent::MouseButtonDblClick; last_press_ms = -100000;
        } else { last_press_ms = now; last_press_w = target; last_press_b = b; last_press_pos = cursor_pos; }
        mouse_btns |= b;
        mouse_capture = target;
    } else if (type == QEvent::MouseButtonRelease) mouse_btns &= ~b;
    /* the mouse dragged away from where the button went down: the next click is not a double click
     * (VB6 oracle f_cool7: a drag that began where the one before began, 0.4 s later) */
    else if (type == QEvent::MouseMove && mouse_btns && (cursor_pos - last_press_pos).manhattanLength() > 4) last_press_ms = -100000;
    QPointF gp(cursor_pos);
    QMouseEvent ev(t, map_from_global(target, gp), map_from_global(target->window(), gp), gp,
                   t == QEvent::MouseMove ? Qt::NoButton : b, mouse_btns, Qt::NoModifier);
    ev.setTimestamp(++mouse_ts);
    QPointer<QWidget> keep(target);
    QCoreApplication::sendEvent(target, &ev);
    if (type == QEvent::MouseButtonRelease && !mouse_btns) mouse_capture = nullptr;
}
static vb_i64 A(vb_dllarg *a) { return a->kind == VBD_DBL ? (vb_i64)a->d : a->i; }
static QWidget *client_of(QWidget *w)
{
    if (auto *fw = dynamic_cast<FormWidget *>(w)) {
        FormState *st = fstate(fw->form);
        if (st->area) return st->area->viewport();
        if (st->canvas) return st->canvas;
    }
    return w;
}
static vb_i64 u_SetCursorPos(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    cursor_pos = QPoint((int)A(&a[0]), (int)A(&a[1]));
    QTimer::singleShot(0, []() { send_mouse(QEvent::MouseMove, Qt::NoButton); });
    return 1;
}
static vb_i64 u_GetCursorPos(vb_i32 argc, vb_dllarg *a)
{
    vb_i32 *p = argc ? (vb_i32 *)vb_dllarg_ptr(&a[0], 8) : nullptr;
    if (!p) return 0;
    p[0] = cursor_pos.x(); p[1] = cursor_pos.y();
    return 1;
}
static vb_i64 map_point(vb_i32 argc, vb_dllarg *a, bool to_screen)
{
    QWidget *w = argc >= 2 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    vb_i32 *p = argc >= 2 ? (vb_i32 *)vb_dllarg_ptr(&a[1], 8) : nullptr;
    if (!w || !p) return 0;
    QWidget *c = client_of(w);
    QPoint r = to_screen ? c->mapToGlobal(QPoint(p[0], p[1])) : c->mapFromGlobal(QPoint(p[0], p[1]));
    p[0] = r.x(); p[1] = r.y();
    return 1;
}
static vb_i64 u_ClientToScreen(vb_i32 argc, vb_dllarg *a) { return map_point(argc, a, true); }
static vb_i64 u_ScreenToClient(vb_i32 argc, vb_dllarg *a) { return map_point(argc, a, false); }
static vb_i64 u_GetWindowRect(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc >= 2 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    vb_i32 *r = argc >= 2 ? (vb_i32 *)vb_dllarg_ptr(&a[1], 16) : nullptr;
    if (!w || !r) return 0;
    if (auto *fw = dynamic_cast<FormWidget *>(w)) {
        FormState *st = fstate(fw->form);
        QPoint tl = client_of(w)->mapToGlobal(QPoint(0, 0));
        Frame f = frame_of(st);
        int mb = st->menubar ? 20 : 0;
        r[0] = tl.x() - f.l; r[1] = tl.y() - (f.t - mb) - mb;
        r[2] = r[0] + (int)((st->clientw + frame_w(st)) / TPP); r[3] = r[1] + (int)((st->clienth + frame_h(st)) / TPP);
        return 1;
    }
    QPoint tl = w->mapToGlobal(QPoint(0, 0));
    r[0] = tl.x(); r[1] = tl.y(); r[2] = tl.x() + w->width(); r[3] = tl.y() + w->height();
    return 1;
}
static vb_i64 u_GetClientRect(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc >= 2 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    vb_i32 *r = argc >= 2 ? (vb_i32 *)vb_dllarg_ptr(&a[1], 16) : nullptr;
    if (!w || !r) return 0;
    r[0] = 0; r[1] = 0;
    if (auto *fw = dynamic_cast<FormWidget *>(w)) {
        FormState *st = fstate(fw->form);
        if (!st->mdi) {                         /* the client area of the form, also in Form_Load (t21 geo: 200x133) */
            r[2] = (int)std::floor(st->clientw / TPP + 1e-6); r[3] = (int)std::floor(st->clienth / TPP + 1e-6);
            return 1;
        }
    }
    QWidget *c = client_of(w);
    r[2] = c->width(); r[3] = c->height();
    return 1;
}
static vb_i64 u_mouse_event(vb_i32 argc, vb_dllarg *a)
{
    long flags = argc > 0 ? (long)A(&a[0]) : 0, dx = argc > 1 ? (long)A(&a[1]) : 0, dy = argc > 2 ? (long)A(&a[2]) : 0;
    /* Windows puts the input in the queue: the program sees it in the message loop (DoEvents) */
    QTimer::singleShot(0, [flags, dx, dy]() {
        if (flags & 1) {                         /* MOUSEEVENTF_MOVE */
            if (flags & 0x8000) {
                QScreen *scr = QGuiApplication::primaryScreen();
                QRect g = scr ? scr->geometry() : QRect(0, 0, 1024, 768);
                cursor_pos = QPoint((int)(dx * g.width() / 65536), (int)(dy * g.height() / 65536));
            } else cursor_pos += QPoint((int)dx, (int)dy);
            send_mouse(QEvent::MouseMove, Qt::NoButton);
        }
        if (flags & 0x2) send_mouse(QEvent::MouseButtonPress, Qt::LeftButton);
        if (flags & 0x4) send_mouse(QEvent::MouseButtonRelease, Qt::LeftButton);
        if (flags & 0x8) send_mouse(QEvent::MouseButtonPress, Qt::RightButton);
        if (flags & 0x10) send_mouse(QEvent::MouseButtonRelease, Qt::RightButton);
        if (flags & 0x20) send_mouse(QEvent::MouseButtonPress, Qt::MiddleButton);
        if (flags & 0x40) send_mouse(QEvent::MouseButtonRelease, Qt::MiddleButton);
    });
    return 0;
}
static vb_i64 u_GetDoubleClickTime(vb_i32, vb_dllarg *) { return QApplication::doubleClickInterval(); }
static vb_i64 u_GetActiveWindow(vb_i32, vb_dllarg *) { return active_form ? hwnd_of(fstate(active_form)->w) : 0; }
/* ---------------------------------------------------------------- GDI and menus of user32 for the Declares
 * A device context is the hWnd of a form or a PictureBox (its hDC); BitBlt and the other GDI calls
 * draw on its picture (the image that AutoRedraw keeps). The desktop is a handle of its own: it can
 * be read (a picture of the screen), drawing on it does nothing. */
static const vb_i32 DESKTOP = 0x7FFE0010;
static Gfx *memdc_gfx(vb_i32 h);                /* a memory DC (winapi_gui.inc) */
static Gfx *gfx_of_dc(vb_i32 h)
{
    if (Gfx *m = memdc_gfx(h)) return m;
    QWidget *w = widget_of_hwnd(h);
    if (!w) return nullptr;
    if (auto *fw = dynamic_cast<FormWidget *>(w)) return &fstate(fw->form)->g;
    Ctl *c = ctl_of_widget(w);
    return c && c->g ? c->g : nullptr;
}
static QImage desktop_image(int x, int y, int w, int h)
{
    QScreen *s = QGuiApplication::primaryScreen();
    QImage img = s ? s->grabWindow(0, x, y, w, h).toImage().convertToFormat(QImage::Format_RGB32) : QImage();
    if (img.isNull()) { img = QImage(w, h, QImage::Format_RGB32); img.fill(Qt::black); }
    return img;
}
static void rop_blt(Gfx *d, int x, int y, const QImage &src, int sx, int sy, int w, int h, vb_u32 rop)
{
    QImage &di = d->img;
    for (int j = 0; j < h; j++) {
        int dy = y + j, syy = sy + j;
        if (dy < 0 || dy >= di.height()) continue;
        QRgb *dl = (QRgb *)di.scanLine(dy);
        const QRgb *sl = syy >= 0 && syy < src.height() ? (const QRgb *)src.constScanLine(syy) : nullptr;
        for (int i = 0; i < w; i++) {
            int dx = x + i, sxx = sx + i;
            if (dx < 0 || dx >= di.width()) continue;
            vb_u32 D = dl[dx] & 0xFFFFFF, S = sl && sxx >= 0 && sxx < src.width() ? sl[sxx] & 0xFFFFFF : 0, r;
            switch (rop) {
            case 0x00EE0086: r = S | D; break;                  /* SRCPAINT */
            case 0x008800C6: r = S & D; break;                  /* SRCAND */
            case 0x00660046: r = S ^ D; break;                  /* SRCINVERT */
            case 0x00440328: r = S & ~D; break;                 /* SRCERASE */
            case 0x00330008: r = ~S; break;                     /* NOTSRCCOPY */
            case 0x001100A6: r = ~(S | D); break;               /* NOTSRCERASE */
            case 0x00BB0226: r = ~S | D; break;                 /* MERGEPAINT */
            case 0x00550009: r = ~D; break;                     /* DSTINVERT */
            case 0x00000042: r = 0; break;                      /* BLACKNESS */
            case 0x00FF0062: r = 0xFFFFFF; break;               /* WHITENESS */
            default: r = S;                                     /* SRCCOPY and the rest */
            }
            dl[dx] = 0xFF000000u | (r & 0xFFFFFF);
        }
    }
    gfx_update(d);
}
static QImage source_image(vb_i32 hsrc, int sx, int sy, int w, int h, int *ox, int *oy)
{
    if (hsrc == DESKTOP || hsrc == 0) { *ox = 0; *oy = 0; return desktop_image(sx, sy, w, h); }
    Gfx *s = gfx_of_dc(hsrc);
    *ox = sx; *oy = sy;
    return s ? s->img : QImage();
}
static vb_i64 g_BitBlt(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 9) return 0;
    Gfx *d = gfx_of_dc((vb_i32)A(&a[0]));
    if (!d) return 1;                           /* the desktop, or a control without a picture */
    int ox, oy, w = (int)A(&a[3]), h = (int)A(&a[4]);
    QImage src = source_image((vb_i32)A(&a[5]), (int)A(&a[6]), (int)A(&a[7]), w, h, &ox, &oy);
    if (src.isNull()) return 0;
    rop_blt(d, (int)A(&a[1]), (int)A(&a[2]), src, ox, oy, w, h, (vb_u32)A(&a[8]));
    return 1;
}
static vb_i64 g_StretchBlt(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 11) return 0;
    Gfx *d = gfx_of_dc((vb_i32)A(&a[0]));
    if (!d) return 1;
    int w = (int)A(&a[3]), h = (int)A(&a[4]), sw = (int)A(&a[8]), sh = (int)A(&a[9]), ox, oy;
    QImage src = source_image((vb_i32)A(&a[5]), (int)A(&a[6]), (int)A(&a[7]), sw, sh, &ox, &oy);
    if (src.isNull() || !sw || !sh || !w || !h) return 0;
    QImage part = src.copy(ox, oy, std::abs(sw), std::abs(sh))
                     .mirrored((sw < 0) != (w < 0), (sh < 0) != (h < 0))
                     .scaled(std::abs(w), std::abs(h), Qt::IgnoreAspectRatio, Qt::FastTransformation);
    rop_blt(d, (int)A(&a[1]) - (w < 0 ? std::abs(w) : 0), (int)A(&a[2]) - (h < 0 ? std::abs(h) : 0),
            part, 0, 0, part.width(), part.height(), (vb_u32)A(&a[10]));
    return 1;
}
static vb_i64 g_GetPixel(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return -1;
    vb_i32 h = (vb_i32)A(&a[0]);
    int x = (int)A(&a[1]), y = (int)A(&a[2]);
    if (h == DESKTOP || h == 0) { QImage i = desktop_image(x, y, 1, 1); return (vb_i64)vbcolor(i.pixel(0, 0)); }
    Gfx *g = gfx_of_dc(h);
    if (!g || x < 0 || y < 0 || x >= g->img.width() || y >= g->img.height()) return 0xFFFFFFFF;   /* CLR_INVALID */
    return (vb_i64)vbcolor(g->img.pixel(x, y));
}
static vb_i64 g_SetPixel(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return -1;
    Gfx *g = gfx_of_dc((vb_i32)A(&a[0]));
    if (!g) return -1;
    long c = (long)A(&a[3]);
    put_px(g, (int)A(&a[1]), (int)A(&a[2]), qcolor(c & 0xFFFFFF));
    gfx_update(g);
    return c & 0xFFFFFF;
}
static vb_i64 g_GetDeviceCaps(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    ensure_app();
    QScreen *s = QGuiApplication::primaryScreen();
    QRect r = s ? s->geometry() : QRect(0, 0, 1024, 768);
    switch ((int)A(&a[1])) {
    case 4: return s ? (vb_i64)s->physicalSize().width() : 320;     /* HORZSIZE (mm) */
    case 6: return s ? (vb_i64)s->physicalSize().height() : 240;    /* VERTSIZE */
    case 8: case 118: return r.width();                         /* HORZRES, DESKTOPHORZRES */
    case 10: case 117: return r.height();                       /* VERTRES, DESKTOPVERTRES */
    case 12: return 32;                                         /* BITSPIXEL */
    case 14: return 1;                                          /* PLANES */
    case 24: return -1;                                         /* NUMCOLORS: more than 256 */
    case 38: return 0x7E99;                                     /* RASTERCAPS */
    case 88: case 90: return 96;                                /* LOGPIXELSX/Y */
    case 104: return 0;                                         /* SIZEPALETTE */
    case 108: return 24;                                        /* COLORRES */
    case 116: return 60;                                        /* VREFRESH */
    }
    return 0;
}
static vb_i64 u_GetSystemMetrics(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 1) return 0;
    ensure_app();
    QScreen *s = QGuiApplication::primaryScreen();
    QRect r = s ? s->geometry() : QRect(0, 0, 1024, 768), av = s ? s->availableGeometry() : r;
    switch ((int)A(&a[0])) {
    case 0: return r.width(); case 1: return r.height();
    case 2: case 3: case 9: case 10: case 20: case 21: return 17;
    case 4: case 31: case 51: return 22;
    case 5: case 6: return 1;
    case 7: case 8: return 3;
    case 11: case 12: case 13: case 14: return 32;
    case 15: return 20;
    case 16: return av.width(); case 17: return av.height() - 22;
    case 19: return 1;                                          /* SM_MOUSEPRESENT */
    case 28: return 136; case 29: return 39;
    case 30: return 36;
    case 32: case 33: return 8;
    case 34: return 136; case 35: return 39;
    case 36: case 37: return 4;
    case 43: return 3;                                          /* SM_CMOUSEBUTTONS */
    case 49: case 50: return 16;
    case 63: return 1;                                          /* SM_NETWORK */
    case 75: return 1;                                          /* SM_MOUSEWHEELPRESENT */
    case 76: case 77: return 0;
    case 78: return QGuiApplication::primaryScreen() ? s->virtualGeometry().width() : r.width();
    case 79: return QGuiApplication::primaryScreen() ? s->virtualGeometry().height() : r.height();
    case 80: return QGuiApplication::screens().size();
    }
    return 0;
}
static vb_i64 u_GetDesktopWindow(vb_i32, vb_dllarg *) { return DESKTOP; }
static vb_i64 u_GetDC(vb_i32 argc, vb_dllarg *a) { vb_i32 h = argc ? (vb_i32)A(&a[0]) : 0; return h ? h : DESKTOP; }
static vb_i64 u_ReleaseDC(vb_i32, vb_dllarg *) { return 1; }
static vb_i64 u_MessageBeep(vb_i32, vb_dllarg *) { vb_beep(); return 1; }
/* the largest top-level window Windows makes (SM_CXMAXTRACK: the screen and a frame): a size in twips
 * given as pixels would ask for a picture of hundreds of megabytes */
static QSize max_window()
{
    QRect scr = QGuiApplication::primaryScreen() ? QGuiApplication::primaryScreen()->virtualGeometry() : QRect(0, 0, 1280, 1024);
    return QSize(scr.width() + 24, scr.height() + 24);
}
static vb_i64 u_SetWindowPos(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 7) return 0;
    QWidget *w = widget_of_hwnd((vb_i32)A(&a[0]));
    if (!w) return 0;
    w = w->window();
    vb_i64 after = A(&a[1]);
    int flags = (int)A(&a[6]);
    if (!(flags & 4) && (after == -1 || after == -2) &&           /* HWND_TOPMOST / HWND_NOTOPMOST */
        w->windowFlags().testFlag(Qt::WindowStaysOnTopHint) != (after == -1)) {
        /* (Qt makes the window anew: later, not inside a Paint event that is drawing it) */
        QPointer<QWidget> pw(w);
        bool top = after == -1;
        QTimer::singleShot(0, w, [pw, top]() {
            if (!pw) return;
            bool vis = pw->isVisible();
            pw->setWindowFlag(Qt::WindowStaysOnTopHint, top);
            if (vis) pw->show();
        });
    }
    if (!(flags & 2)) w->move((int)A(&a[2]), (int)A(&a[3]));    /* SWP_NOMOVE */
    if (!(flags & 1)) {                                         /* SWP_NOSIZE: the outer size */
        QRect fg = w->frameGeometry(), g = w->geometry();
        QSize mx = max_window();                                /* (Windows: no bigger than the screen + frame) */
        int cx = std::min((int)A(&a[4]), mx.width()), cy = std::min((int)A(&a[5]), mx.height());
        w->resize(std::max(1, cx - (fg.width() - g.width())), std::max(1, cy - (fg.height() - g.height())));
    }
    if (flags & 0x80) w->hide();
    if (flags & 0x40) w->show();
    return 1;
}
/* menus: GetMenu (the bar of a form), GetSubMenu, GetMenuItemCount, TrackPopupMenu */
static std::map<vb_i32, QPointer<QWidget>> menu_handles;
static vb_i32 menu_handle(QWidget *m)
{
    if (!m) return 0;
    for (auto &e : menu_handles) if (e.second == m) return e.first;
    static vb_i32 next = 0x00090010;
    vb_i32 h = next; next += 4;
    menu_handles[h] = m;
    return h;
}
static QList<QAction *> visible_actions(QWidget *m)
{
    QList<QAction *> r;
    if (m) for (QAction *ac : m->actions()) if (ac->isVisible() && !ac->isSeparator()) r.push_back(ac);
    return r;
}
static vb_i64 u_GetMenu(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    auto *fw = w ? dynamic_cast<FormWidget *>(w->window()) : nullptr;
    return fw && fstate(fw->form)->menubar ? menu_handle(fstate(fw->form)->menubar) : 0;
}
static vb_i64 u_GetSubMenu(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    auto it = menu_handles.find((vb_i32)A(&a[0]));
    if (it == menu_handles.end() || !it->second) return 0;
    QList<QAction *> acts = visible_actions(it->second);
    int pos = (int)A(&a[1]);
    if (pos < 0 || pos >= acts.size() || !acts[pos]->menu()) return 0;
    return menu_handle(acts[pos]->menu());
}
static vb_i64 u_GetMenuItemCount(vb_i32 argc, vb_dllarg *a)
{
    auto it = argc ? menu_handles.find((vb_i32)A(&a[0])) : menu_handles.end();
    return it == menu_handles.end() || !it->second ? -1 : visible_actions(it->second).size();
}
static vb_i64 u_TrackPopupMenu(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    auto it = menu_handles.find((vb_i32)A(&a[0]));
    auto *m = it == menu_handles.end() ? nullptr : qobject_cast<QMenu *>(it->second.data());
    if (!m) return 0;
    flush_pending();
    m->exec(QPoint((int)A(&a[2]), (int)A(&a[3])));             /* the Click of the item runs inside */
    return 1;
}

/* SetTimer / KillTimer with a TimerProc (AddressOf): the runtime's timers call the procedure */
static std::map<vb_i32, QTimer *> api_timers;
static vb_i64 u_SetTimer(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    vb_i32 hwnd = (vb_i32)A(&a[0]), id = (vb_i32)A(&a[1]), ms = (vb_i32)A(&a[2]);
    vb_i64 proc = A(&a[3]);
    if (!vb_callback_fn(proc)) return 0;        /* a TimerProc that is no AddressOf of the program */
    static vb_i32 next_id = 1;
    if (!hwnd || !id) id = 0x4000 + next_id++;   /* SetTimer(0, 0, ...): a new timer number */
    auto old = api_timers.find(id);
    if (old != api_timers.end()) { old->second->stop(); old->second->deleteLater(); api_timers.erase(old); }
    auto *t = new QTimer();
    t->setInterval(std::max(1, ms));
    QObject::connect(t, &QTimer::timeout, [hwnd, id, proc]() {
        vb_i32 args[4] = { hwnd, 0x113, id, (vb_i32)QDateTime::currentMSecsSinceEpoch() };   /* WM_TIMER */
        vb_callback(proc, 4, args);
        if (vb_errpend) vb_unhandled();
    });
    t->start();
    api_timers[id] = t;
    return id;
}
static vb_i64 u_KillTimer(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    auto it = api_timers.find((vb_i32)A(&a[1]));
    if (it == api_timers.end()) return 0;
    it->second->stop();
    it->second->deleteLater();
    api_timers.erase(it);
    return 1;
}

#include "winapi_gui.inc"

/* AppActivate title: a form of the program whose caption is the title (or starts with it) */
static void appactivate_hook(vbs title)
{
    QString t = qs(title);
    for (vbobj f : loaded_forms) {
        QWidget *w = fstate(f)->w;
        if (w && (w->windowTitle().compare(t, Qt::CaseInsensitive) == 0 || w->windowTitle().startsWith(t, Qt::CaseInsensitive))) {
            w->raise(); w->activateWindow();
            return;
        }
    }
}

static vb_apiext api_hook(const char *lib, const char *name)
{
    if (vb_apiext f = gui_api(lib, name)) return f;
    if (!strcmp(lib, "gdi32")) {
        static const struct { const char *n; vb_apiext f; } t[] = {
            { "BitBlt", g_BitBlt }, { "StretchBlt", g_StretchBlt }, { "GetPixel", g_GetPixel },
            { "SetPixel", g_SetPixel }, { "SetPixelV", g_SetPixel }, { "GetDeviceCaps", g_GetDeviceCaps } };
        for (auto &e : t) if (!strcmp(e.n, name)) return e.f;
        return nullptr;
    }
    if (strcmp(lib, "user32")) return nullptr;
    static const struct { const char *n; vb_apiext f; } t2[] = {
        { "GetSystemMetrics", u_GetSystemMetrics }, { "GetDesktopWindow", u_GetDesktopWindow },
        { "GetDC", u_GetDC }, { "GetWindowDC", u_GetDC }, { "ReleaseDC", u_ReleaseDC },
        { "MessageBeep", u_MessageBeep }, { "SetWindowPos", u_SetWindowPos }, { "GetMenu", u_GetMenu },
        { "GetSubMenu", u_GetSubMenu }, { "GetMenuItemCount", u_GetMenuItemCount },
        { "TrackPopupMenu", u_TrackPopupMenu }, { "SetTimer", u_SetTimer }, { "KillTimer", u_KillTimer } };
    for (auto &e : t2) if (!strcmp(e.n, name)) return e.f;
    static const struct { const char *n; vb_apiext f; } t[] = {
        { "SetCursorPos", u_SetCursorPos }, { "GetCursorPos", u_GetCursorPos }, { "ClientToScreen", u_ClientToScreen },
        { "ScreenToClient", u_ScreenToClient }, { "GetWindowRect", u_GetWindowRect }, { "GetClientRect", u_GetClientRect },
        { "mouse_event", u_mouse_event }, { "GetDoubleClickTime", u_GetDoubleClickTime },
        { "GetActiveWindow", u_GetActiveWindow }, { "GetForegroundWindow", u_GetActiveWindow } };
    for (auto &e : t) if (!strcmp(e.n, name)) return e.f;
    return nullptr;
}

extern "C" vbobj vbf_forms(void)
{
    return make_coll(loaded_forms);
}

/* the user moves the focus (mouse, Tab): VB6 events as with SetFocus */
static void on_focus_changed(QWidget *, QWidget *now)
{
    if (!now) return;
    FormState *st = form_of_widget(now);
    if (!st || !st->loaded) return;
    Ctl *c = ctl_of_widget(now);
    if (c && c->form && c->form != st && c->form->host) {   /* a control in a user control: the user control */
        Ctl *inner = c;
        Ctl *h = c->form->host;
        if (h->form != st) return;
        if (h->self == st->focus) {             /* from one of its controls to another */
            FormState *us = c->form;
            vbobj was = us->focus;
            if (was == inner->self) return;
            us->focus = inner->self;
            vbobj now_in = inner->self;
            if (was) vb_oaddref(was);
            vb_oaddref(now_in);
            post([was, now_in]() {
                if (was) { Ctl *w = C(was); if (w && w->form) fire_ctl(w, "lostfocus"); vb_orelease(was); }
                Ctl *n = C(now_in); if (n && n->form) fire_ctl(n, "gotfocus");
                vb_orelease(now_in);
            });
            return;
        }
        c->form->focus = inner->self;           /* the one the user chose */
        focus_to(st, h, true, true);
        return;
    }
    if (c && c->form == st && c->self != st->focus) focus_to(st, c, true, true);
}

static void install_filters()
{
    static KeyFilter filter;
    app->installEventFilter(&filter);
    QObject::connect(app, &QApplication::focusChanged, on_focus_changed);
}

static void monkey_start();
static void doevents_hook(void)
{
    /* the robot: also a program that never gets to the event loop (Do ... DoEvents ... Loop in Form_Load) */
    if (app && monkey_on && !loaded_forms.empty()) monkey_start();
    if (app) QCoreApplication::processEvents(QEventLoop::AllEvents, 50);
    flush_pending();
}

extern "C" void vbf_start(vbobj form) { form_show(form, 0); }

/* a program without forms that uses Screen, Clipboard or pictures: links this file (and its hooks) */
extern "C" void vbf_use(void) {}

/* ---------------------------------------------------------------- the test robot (VBRT_MONKEY=1)
 * LinBASIC Studio's own check of a program: it clicks every button, option, check box and menu
 * item of the forms that are open, one by one (the modal form first, the buttons that end the
 * program never), answers MsgBox with the first button, and writes every action and every
 * run-time error to stderr - an error does not end the program. */
std::vector<vbobj> modal_stack;
static QString monkey_doing;
static void monkey_log(const QString &what)
{
    fprintf(stderr, "monkey: %s\n", what.toUtf8().constData());
    fflush(stderr);
}
static int monkey_error(vb_i32 n, const char *d)
{
    if (!monkey_on) return 0;
    monkey_log(QString("ERROR %1 (%2) at %3 during %4").arg(n).arg(QString::fromUtf8(d)).arg(QString::fromUtf8(vb_where())).arg(monkey_doing));
    return 1;
}
static int monkey_compile_error(const char *m)
{
    if (!monkey_on) return 0;
    monkey_log(QString("COMPILE ERROR (%1) at %2 during %3").arg(QString::fromUtf8(m)).arg(QString::fromUtf8(vb_where())).arg(monkey_doing));
    return 1;
}
static QString plain_caption(const QString &c) { QString t = c; t.remove('&'); return t.trimmed().toLower().remove('.'); }
static bool ends_program(const QString &cap)
{
    static const QStringList w = { "exit", "quit", "end", "e&xit" };
    return w.contains(cap) || cap.startsWith("exit ") || cap.startsWith("quit ");
}
static bool dismisses(const QString &cap)
{
    static const QStringList w = { "ok", "cancel", "close", "done", "return", "back" };
    return w.contains(cap);
}
/* what was done: Form.Control(index) and its caption - a button whose caption changed (Start / Stop) is
 * new; a second instance of a form (File > New) is not */
static std::set<std::string> monkey_done;
/* the action runs after the timer's slot has returned: a Click that shows a modal form or loops
 * with DoEvents does not stop the robot (Qt does not run a timer's slot inside itself) */
static void monkey_later(std::function<void()> f) { QTimer::singleShot(0, app, std::move(f)); }
/* a menu item can be clicked when every menu above it is visible and enabled */
static bool menu_reachable(FormState *st, Ctl *c)
{
    for (int depth = 0; depth < 10; depth++) {
        Ctl *up = nullptr;
        for (vbobj o : st->ctls) {
            Ctl *m = C(o);
            if (m && m->kind == K_MENU && m->menu && m->menu->actions().contains(c->act)) { up = m; break; }
        }
        if (!up) return true;
        if (!up->enabled || !up->visible) return false;
        c = up;
    }
    return true;
}
static int monkey_idle;
/* a control array: the robot tries its first elements only (a grid of 300 buttons kept the robot busy past its
 * 45 seconds and the program counted as HANG: minesweeper, sudoku, dynamic buttons) */
static const int MONKEY_ARRAY_MAX = 6;
static std::map<std::string, int> monkey_arr;
static void monkey_took(const QString &fname, Ctl *c)
{
    if (c->index >= 0) monkey_arr[(fname + "." + c->name).toStdString()]++;
}
static bool monkey_try(vbobj form, bool dismiss_ok, bool main_form)
{
    FormState *st = fstate(form);
    if (!st->loaded) return false;
    QString fname = QString::fromUtf8(formdef(form)->name);
    for (vbobj co : st->ctls) {
        Ctl *c = C(co);
        if (!c || !c->enabled || !c->visible) continue;
        if (!st->visible && c->kind != K_MENU) continue;    /* a hidden form: its menus (PopupMenu) */
        QString cap = plain_caption(c->caption), lname = c->name.toLower();
        if (ends_program(cap) || lname.contains("exit") || lname.contains("quit")) continue;
        bool dism = dismisses(cap);
        if (dism != dismiss_ok) continue;
        if (dism && main_form) continue;
        QString who = fname + "." + c->name + (c->index >= 0 ? QString("(%1)").arg(c->index) : QString());
        std::string key = (who + "|" + c->caption).toStdString();
        if (monkey_done.count(key)) continue;
        if (c->index >= 0 && monkey_arr[(fname + "." + c->name).toStdString()] >= MONKEY_ARRAY_MAX) continue;
        if (c->kind == K_MENU) {
            if (!c->act || c->menu || !c->act->isEnabled() || cap == "-" || cap.isEmpty() || !menu_reachable(st, c)) continue;
            monkey_done.insert(key); monkey_took(fname, c);
            monkey_doing = who + " [menu " + c->caption + "]";
            monkey_log("click " + monkey_doing);
            QPointer<QAction> a = c->act;
            monkey_later([a] { if (a) a->trigger(); });
            return true;
        }
        if (c->kind == K_COMMAND || c->kind == K_CHECK || c->kind == K_OPTION) {
            auto *b = qobject_cast<QAbstractButton *>(c->w.data());
            if (!b || !b->isVisibleTo(st->w)) continue;
            monkey_done.insert(key); monkey_took(fname, c);
            monkey_doing = who + " [" + c->caption + "]";
            monkey_log("click " + monkey_doing);
            QPointer<QAbstractButton> pb = b;
            monkey_later([pb] { if (pb) pb->click(); });
            return true;
        }
        if (c->kind == K_TEXT && c->w && c->w->isVisibleTo(st->w) && get_text(c).isEmpty()) {
            auto *le = qobject_cast<QLineEdit *>(c->w.data());
            auto *pe = qobject_cast<QPlainTextEdit *>(c->w.data());
            if ((le && le->isReadOnly()) || (pe && pe->isReadOnly())) continue;
            monkey_done.insert(key); monkey_took(fname, c);            /* an empty text box gets "1": buttons that wait for text */
            monkey_doing = who + " [type 1]";
            monkey_log("type " + monkey_doing);
            vbobj self = c->self;
            monkey_later([self] { Ctl *x = C(self); if (x && x->form) { set_text(x, "1", true); } });
            return true;
        }
        if (c->kind == K_OCX && c->w && c->w->isVisibleTo(st->w)) {
            /* the parts of an .ocx control the user clicks: Toolbar buttons, tabs, the first nodes and items */
            std::vector<std::pair<QString, std::function<void()>>> parts;
            for (auto *tb : c->w->findChildren<QToolButton *>()) {
                if (!tb->isVisible() || !tb->isEnabled()) continue;
                QString t = tb->text().isEmpty() ? tb->toolTip() : tb->text();
                QPointer<QToolButton> p = tb;
                parts.push_back({"button " + t + QString::number(parts.size()), [p] { if (p) p->click(); }});
            }
            for (auto *bar : c->w->findChildren<QTabBar *>()) {
                QPointer<QTabBar> p = bar;
                for (int i = 0; i < bar->count(); i++)
                    if (bar->isTabEnabled(i)) parts.push_back({"tab " + bar->tabText(i), [p, i] { if (p) p->setCurrentIndex(i); }});
            }
            for (auto *tw : c->w->findChildren<QTreeWidget *>()) {
                if (!tw->isVisible()) continue;
                QPointer<QTreeWidget> p = tw;
                for (int i = 0; i < std::min(5, tw->topLevelItemCount()); i++)
                    parts.push_back({"item " + tw->topLevelItem(i)->text(0), [p, i] {
                        if (!p || i >= p->topLevelItemCount()) return;
                        QTreeWidgetItem *it = p->topLevelItem(i); p->setCurrentItem(it); emit p->itemClicked(it, 0); }});
            }
            for (auto *lw : c->w->findChildren<QListWidget *>()) {
                if (!lw->isVisible()) continue;
                QPointer<QListWidget> p = lw;
                for (int i = 0; i < std::min(5, lw->count()); i++)
                    parts.push_back({"item " + lw->item(i)->text(), [p, i] {
                        if (!p || i >= p->count()) return;
                        p->setCurrentRow(i); emit p->itemClicked(p->item(i)); }});
            }
            for (auto &pt : parts) {
                std::string k = (who + "|" + pt.first).toStdString();
                if (monkey_done.count(k)) continue;
                monkey_done.insert(k); monkey_took(fname, c);
                monkey_doing = who + " [" + pt.first + "]";
                monkey_log("click " + monkey_doing);
                monkey_later(pt.second);
                return true;
            }
            continue;
        }
        if ((c->kind == K_LIST || c->kind == K_COMBO || c->kind == K_FILE || c->kind == K_DIR) && list_count(c) > 0) {
            monkey_done.insert(key); monkey_took(fname, c);
            int i = std::min(1, list_count(c) - 1);
            monkey_doing = who + QString(" [select %1]").arg(i);
            monkey_log("click " + monkey_doing);
            QPointer<QWidget> w = c->w;
            Ctl *cc = c;
            monkey_later([w, cc, i] {
                if (!w) return;
                if (auto *lw = LW(cc)) { lw->setCurrentRow(i); if (lw->item(i)) emit lw->itemClicked(lw->item(i)); }
                else if (auto *cb = CB(cc)) { cb->setCurrentIndex(i); emit cb->activated(i); }
            });
            return true;
        }
    }
    return false;
}
/* pictures of the forms that are open (VBRT_MONKEY_SHOTS=folder) */
static void monkey_shots(const char *when)
{
    const char *dir = getenv("VBRT_MONKEY_SHOTS");
    if (!dir) return;
    for (vbobj f : loaded_forms) {
        FormState *st = fstate(f);
        if (!st->visible || !st->w) continue;
        QWidget *top = st->w->window();
        top->grab().save(QString("%1/%2-%3.png").arg(QString::fromUtf8(dir), QString::fromUtf8(formdef(f)->name), when));
    }
}
static void monkey_step()
{
    if (!monkey_on) return;
    static int steps;
    if (++steps == 3) monkey_shots("start");
    std::vector<vbobj> forms;
    if (!modal_stack.empty()) forms.push_back(modal_stack.back());
    else for (auto it = loaded_forms.rbegin(); it != loaded_forms.rend(); ++it) forms.push_back(*it);
    vbobj main_form = loaded_forms.empty() ? nullptr : loaded_forms.front();
    for (vbobj f : forms) if (monkey_try(f, false, f == main_form)) { monkey_idle = 0; return; }
    for (vbobj f : forms) if (f != main_form && monkey_try(f, true, false)) { monkey_idle = 0; return; }
    if (!modal_stack.empty()) {                 /* a modal form with nothing more to do: close it */
        vbobj f = modal_stack.back();
        /* a form that will not unload (Cancel in QueryUnload, or it loads itself again) kept the robot unloading
         * it until the 45 seconds were over (robot: super error handler, puzzle game): Hide it (that ends a modal
         * Show too), and at last the end of the run */
        static vbobj last; static int tries;
        tries = (f == last) ? tries + 1 : 1; last = f;
        if (tries > 6) {
            monkey_log("modal form " + QString::fromUtf8(formdef(f)->name) + " does not close: done");
            monkey_shots("end");
            monkey_log("done");
            fflush(stdout);
            std::_Exit(0);
        }
        if (tries > 3) {
            monkey_log("hide " + QString::fromUtf8(formdef(f)->name));
            monkey_later([f] { form_hide(f); });
            return;
        }
        monkey_log("unload " + QString::fromUtf8(formdef(f)->name));
        monkey_later([f] { unload_form(f, 0); });
        return;
    }
    if (++monkey_idle >= 8) {
        monkey_shots("end");
        monkey_log("done");
        fflush(stdout);
        std::_Exit(0);
    }
}
static void monkey_start()
{
    static bool started;
    if (started) return;                        /* (from vbf_run, or from a DoEvents loop in Form_Load) */
    started = true;
    monkey_on = true;
    vb_unhandled_hook = monkey_error;
    auto *t = new QTimer(app);
    QObject::connect(t, &QTimer::timeout, monkey_step);
    t->start(150);
}

/* a program started on its own (not in the IDE, not the robot): VB6's error box, then the end */
static int error_box(vb_i32 n, const char *d)
{
    if (!app || getenv("VBRT_NO_ERRBOX") || getenv("VBRT_DEBUG")) return 0;
    vb_i32 m = vb_mark();
    QString title = qs(vb_app_title());        /* VB6: App.Title (h1: "H1"), not the name Qt has ("vbapp") */
    vb_release(m);
    QMessageBox::warning(dialog_parent(), title,       /* VB6's made program: the exclamation mark (h1, H-3) */ QString("Run-time error '%1':\n\n%2").arg(n).arg(QString::fromUtf8(d)));
    return 0;
}

static QString first_upper(QString s) { if (!s.isEmpty()) s[0] = s[0].toUpper(); return s; }
/* Compile On Demand: VB6's box "Compile error:" (a --debug program started on its own), then the end */
static int compile_box(const char *m)
{
    if (!app || getenv("VBRT_NO_ERRBOX") || getenv("VBRT_DEBUG")) return 0;
    vb_i32 mk = vb_mark();
    QString title = qs(vb_app_title());
    vb_release(mk);
    QMessageBox::critical(dialog_parent(), title, QString("Compile error:\n\n%1").arg(first_upper(QString::fromUtf8(m))));
    return 0;
}

/* VBRT_SCREENSHOT=file.png: a picture of the program's forms after VBRT_SCREENSHOT_DELAY ms (default 1500),
 * then the program ends (pictures for documentation; the active form, or each form as file-<Form>.png) */
static void screenshot_start(const char *file)
{
    const char *d = getenv("VBRT_SCREENSHOT_DELAY");
    QString path = QString::fromUtf8(file);
    QTimer::singleShot(d ? atoi(d) : 1500, app, [path]() {
        std::vector<vbobj> shown;
        for (vbobj f : loaded_forms) if (fstate(f)->visible && fstate(f)->w) shown.push_back(f);
        for (vbobj f : shown) {
            QString out = path;
            if (shown.size() > 1 && f != active_form) {
                int dot = out.lastIndexOf('.');
                out.insert(dot < 0 ? out.size() : dot, QString("-") + QString::fromUtf8(formdef(f)->name));
            }
            fstate(f)->w->window()->grab().save(out);
        }
        fflush(stdout);
        std::_Exit(0);
    });
}

extern "C" void vbf_run(void)
{
    vb_doevents_hook = doevents_hook;
    if (loaded_forms.empty() || !app) return;
    if (getenv("VBRT_MONKEY")) monkey_start();
    else if (const char *shot = getenv("VBRT_SCREENSHOT")) screenshot_start(shot);
    running = true;
    app->exec();
    running = false;
}

#include "ocx.inc"
#include "uctl.inc"
#include "uctl_async.inc"
#include "data.inc"
#include "report.inc"
#include "ocx_data.inc"
#include "ocx_rdc.inc"
#include "ocx_drep.inc"
#include "help.inc"

static struct Init {
    Init()
    {
        vb_doevents_hook = doevents_hook; vb_sendkeys_hook = sendkeys_hook; vb_appactivate_hook = appactivate_hook;
        vb_loadpicture_hook = loadpicture_hook; vb_savepicture_hook = savepicture_hook; vb_picture_data_hook = picture_data_hook;
        vb_msgbox_hook = msgbox_hook; vb_inputbox_hook = inputbox_hook;
        vb_api_hook = api_hook;
        vb_ds_changed_ctls_hook = ds_changed_ctls;
        vb_obj_inner = uctl_inner;
        monkey_on = getenv("VBRT_MONKEY") != nullptr;
        vb_unhandled_hook = monkey_on ? monkey_error : error_box;
        vb_compile_hook = monkey_on ? monkey_compile_error : compile_box;
    }
} init_hook;
