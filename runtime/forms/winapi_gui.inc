/* winapi_gui.inc - the Windows procedures of windows, keys, messages, drawing and dialogs for the
 * Declares of programs (part of vbforms.cpp, with u_GetCursorPos ... before it): user32 (windows by
 * hWnd, SendMessage for the controls, GetAsyncKeyState, window styles and layered windows, regions
 * with SetWindowRgn), gdi32 (pens, brushes, memory DCs and bitmaps, lines, shapes, text, regions),
 * shell32 (Shell_NotifyIcon in the tray, SHBrowseForFolder), comdlg32 (GetOpenFileName,
 * GetSaveFileName, ChooseColor), msimg32 (TransparentBlt, AlphaBlend), oleaut32 (OleTranslateColor).
 * On Linux a window outside the program cannot be found or changed: FindWindow finds the program's
 * own forms, the rest answers as Windows does when a window is not there.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <QSystemTrayIcon>
#include <QWindow>
#include <QDesktopServices>
#include <QBitmap>
#include <QColorDialog>
#include <QPainterPath>
#include <QCoreApplication>

extern "C" {
#ifndef _WIN32
vb_i32 vb_pidl_make(const char *path);
#endif
vbs vb_sfromansi(const char *s, int n);
char *vb_stoansi(vbs s, int *n);
}

/* ---------------------------------------------------------------- arguments */
static QString astr(vb_dllarg *a)
{
    char *s = vb_dllarg_str(a);
    if (!s) return QString();
    vb_i32 mk = vb_mark();
    QString r = qs(vb_sfromansi(s, (int)strlen(s)));
    vb_release(mk);
    return r;
}
/* text into a buffer of the program (size n); the length written */
static vb_i64 put_astr(vb_dllarg *a, vb_i64 n, const QString &v)
{
    char *b = vb_dllarg_str(a);
    vb_i32 mk = vb_mark();
    int an = 0;
    char *t = vb_stoansi(vs(v), &an);
    if (b && n > 0) {
        int k = an < n - 1 ? an : (int)n - 1;
        memcpy(b, t, (size_t)k); b[k] = 0;
        an = k;
    }
    vb_release(mk);
    return an;
}
static vbobj obj_of_hwnd(vb_i32 h)
{
    QWidget *w = widget_of_hwnd(h);
    if (!w) return nullptr;
    if (auto *fw = dynamic_cast<FormWidget *>(w)) return fw->form;
    Ctl *c = ctl_of_widget(w);
    return c ? c->self : nullptr;
}
static QWidget *top_of_hwnd(vb_i32 h) { QWidget *w = widget_of_hwnd(h); return w ? w->window() : nullptr; }

/* ---------------------------------------------------------------- keys: GetAsyncKeyState, GetKeyState */
static bool vk_state[256], vk_hit[256];
static void key_track(QKeyEvent *ke, bool down)
{
    int vk = vk_of_qt(ke->key(), ke->text());
    if (vk <= 0 || vk > 255) return;
    if (down && !vk_state[vk]) vk_hit[vk] = true;
    vk_state[vk] = down;
}
static bool vk_is_down(int vk)
{
    Qt::MouseButtons mb = QGuiApplication::mouseButtons();
    Qt::KeyboardModifiers km = QGuiApplication::queryKeyboardModifiers();
    switch (vk) {
    case 1: return mb & Qt::LeftButton;
    case 2: return mb & Qt::RightButton;
    case 4: return mb & Qt::MiddleButton;
    case 0x10: case 0xA0: case 0xA1: return km & Qt::ShiftModifier;
    case 0x11: case 0xA2: case 0xA3: return km & Qt::ControlModifier;
    case 0x12: case 0xA4: case 0xA5: return km & Qt::AltModifier;
    }
    return vk > 0 && vk < 256 && vk_state[vk];
}
static vb_i64 u_GetAsyncKeyState(vb_i32 argc, vb_dllarg *a)
{
    int vk = argc ? (int)A(&a[0]) & 0xFF : 0;
    bool hit = vk_hit[vk];
    vk_hit[vk] = false;
    return (vk_is_down(vk) ? -32768 : 0) | (hit ? 1 : 0);
}
static vb_i64 u_GetKeyState(vb_i32 argc, vb_dllarg *a)
{
    int vk = argc ? (int)A(&a[0]) & 0xFF : 0;
    return vk_is_down(vk) ? -127 : 0;           /* 0xFF81: down (the toggle bit is not known here) */
}
static vb_i64 u_GetKeyboardState(vb_i32 argc, vb_dllarg *a)
{
    unsigned char *b = argc ? (unsigned char *)vb_dllarg_ptr(&a[0], 256) : nullptr;
    if (!b) return 0;
    for (int i = 0; i < 256; i++) b[i] = vk_is_down(i) ? 0x80 : 0;
    return 1;
}
static vb_i64 u_keybd_event(vb_i32 argc, vb_dllarg *a)
{
    int vk = argc ? (int)A(&a[0]) & 0xFF : 0;
    long flags = argc > 2 ? (long)A(&a[2]) : 0;
    QTimer::singleShot(0, [vk, flags]() {
        QWidget *f = QApplication::focusWidget();
        if (!f) return;
        int key = vk >= 'A' && vk <= 'Z' ? Qt::Key_A + (vk - 'A') : vk >= '0' && vk <= '9' ? Qt::Key_0 + (vk - '0') :
                  vk == 13 ? Qt::Key_Return : vk == 8 ? Qt::Key_Backspace : vk == 9 ? Qt::Key_Tab : vk == 27 ? Qt::Key_Escape :
                  vk == 32 ? Qt::Key_Space : vk == 0x10 ? Qt::Key_Shift : vk == 0x11 ? Qt::Key_Control : vk == 0x12 ? Qt::Key_Alt :
                  vk >= 0x25 && vk <= 0x28 ? Qt::Key_Left + (vk - 0x25) : vk >= 0x70 && vk <= 0x7F ? Qt::Key_F1 + (vk - 0x70) : 0;
        if (!key) return;
        QString text = (vk >= 'A' && vk <= 'Z') ? QString(QChar(vk + 32)) : (vk >= '0' && vk <= '9') || vk == 32 ? QString(QChar(vk)) : QString();
        QKeyEvent ev(flags & 2 ? QEvent::KeyRelease : QEvent::KeyPress, key, Qt::NoModifier, flags & 2 ? QString() : text);
        QCoreApplication::sendEvent(f, &ev);
    });
    return 0;
}

/* ---------------------------------------------------------------- windows */
static vb_i64 u_FindWindowA(vb_i32 argc, vb_dllarg *a)
{
    QString cls = argc > 0 ? astr(&a[0]) : QString(), title = argc > 1 ? astr(&a[1]) : QString();
    if (title.isEmpty() && !cls.isEmpty() && !cls.startsWith("Thunder", Qt::CaseInsensitive)) return 0;
    for (vbobj f : loaded_forms) {
        FormState *st = fstate(f);
        if (!st->w) continue;
        if (title.isEmpty() || st->w->windowTitle().compare(title, Qt::CaseInsensitive) == 0) return hwnd_of(st->w);
    }
    return 0;
}
static vb_i64 u_FindWindowExA(vb_i32 argc, vb_dllarg *a)
{
    QWidget *parent = argc > 0 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    QString title = argc > 3 ? astr(&a[3]) : QString();
    if (!parent) return 0;
    for (QWidget *c : parent->findChildren<QWidget *>()) {
        Ctl *k = ctl_of_widget(c);
        if (!k || k->w != c) continue;
        if (title.isEmpty() || k->caption.compare(title, Qt::CaseInsensitive) == 0) return hwnd_of(c);
    }
    return 0;
}
static QString window_text(QWidget *w)
{
    if (!w) return QString();
    if (dynamic_cast<FormWidget *>(w)) return w->windowTitle();
    if (auto *e = qobject_cast<QLineEdit *>(w)) return e->text();
    if (auto *e = qobject_cast<QPlainTextEdit *>(w)) return e->toPlainText();
    if (auto *e = qobject_cast<QTextEdit *>(w)) return e->toPlainText();
    if (auto *b = qobject_cast<QAbstractButton *>(w)) return b->text().remove('&');
    if (auto *l = qobject_cast<QLabel *>(w)) return l->text();
    if (auto *g = qobject_cast<QGroupBox *>(w)) return g->title();
    Ctl *c = ctl_of_widget(w);
    return c ? c->caption : w->windowTitle();
}
static vb_i64 u_GetWindowTextA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    return put_astr(&a[1], A(&a[2]), window_text(widget_of_hwnd((vb_i32)A(&a[0]))));
}
static vb_i64 u_GetWindowTextLengthA(vb_i32 argc, vb_dllarg *a)
{
    return argc ? window_text(widget_of_hwnd((vb_i32)A(&a[0]))).size() : 0;
}
static vb_i64 u_SetWindowTextA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    vbobj o = obj_of_hwnd((vb_i32)A(&a[0]));
    if (!o) return 0;
    QWidget *w = widget_of_hwnd((vb_i32)A(&a[0]));
    vb_i32 mk = vb_mark();
    vbvar v = vb_v_s(vs(astr(&a[1])));
    bool text = qobject_cast<QLineEdit *>(w) || qobject_cast<QPlainTextEdit *>(w) || qobject_cast<QTextEdit *>(w);
    vb_late_o(o, text ? "text" : "caption", VBM_LET, 1, &v);
    if (vb_errpend) vb_err_clear_pending();
    vb_release(mk);
    return 1;
}
static vb_i64 u_ShowWindow(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w) return 0;
    bool was = w->isVisible();
    int cmd = argc > 1 ? (int)A(&a[1]) : 1;
    QWidget *t = w->isWindow() ? w : w;
    switch (cmd) {
    case 0: t->hide(); break;
    case 2: case 6: case 7: case 11: if (t->isWindow()) t->showMinimized(); break;
    case 3: if (t->isWindow()) t->showMaximized(); else t->show(); break;
    case 9: if (t->isWindow()) t->showNormal(); else t->show(); break;
    default: t->show(); break;
    }
    return was;
}
static vb_i64 u_IsWindow(vb_i32 argc, vb_dllarg *a) { return argc && widget_of_hwnd((vb_i32)A(&a[0])) ? 1 : 0; }
static vb_i64 u_IsWindowVisible(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr; return w && w->isVisible(); }
static vb_i64 u_IsWindowEnabled(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr; return w && w->isEnabled(); }
static vb_i64 u_IsIconic(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? top_of_hwnd((vb_i32)A(&a[0])) : nullptr; return w && w->isMinimized(); }
static vb_i64 u_IsZoomed(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? top_of_hwnd((vb_i32)A(&a[0])) : nullptr; return w && w->isMaximized(); }
static vb_i64 u_EnableWindow(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc > 1 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w) return 0;
    bool was = !w->isEnabled();
    w->setEnabled(A(&a[1]) != 0);
    return was;
}
static vb_i64 u_MoveWindow(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc >= 5 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w) return 0;
    int x = (int)A(&a[1]), y = (int)A(&a[2]), cx = (int)A(&a[3]), cy = (int)A(&a[4]);
    if (w->isWindow()) {
        QRect fg = w->frameGeometry(), g = w->geometry();
        QSize mx = max_window();
        cx = std::min(cx, mx.width()); cy = std::min(cy, mx.height());
        w->move(x, y);
        w->resize(std::max(1, cx - (fg.width() - g.width())), std::max(1, cy - (fg.height() - g.height())));
    } else w->setGeometry(x, y, cx, cy);
    return 1;
}
static vb_i64 u_BringWindowToTop(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc ? top_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w) return 0;
    w->raise(); w->activateWindow();
    return 1;
}
static vb_i64 u_GetParent(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w || w->isWindow()) return 0;
    for (QWidget *p = w->parentWidget(); p; p = p->parentWidget()) {
        if (dynamic_cast<FormWidget *>(p)) return hwnd_of(p);
        Ctl *c = ctl_of_widget(p);
        if (c && c->w == p) return hwnd_of(p);
    }
    return 0;
}
static vb_i64 u_SetParent(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc > 1 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    QWidget *p = argc > 1 ? widget_of_hwnd((vb_i32)A(&a[1])) : nullptr;
    if (!w) return 0;
    vb_i64 old = u_GetParent(1, a);
    if (p) {
        QWidget *host = client_of(p);
        bool vis = w->isVisible();
        QPoint at = w->isWindow() ? QPoint(0, 0) : w->pos();
        w->setParent(host);
        w->move(at);
        if (vis) w->show();
    }
    return old ? old : 1;
}
static vb_i64 u_DestroyWindow(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w) return 0;
    if (w->isWindow()) w->close(); else w->hide();
    return 1;
}
static vb_i64 u_CloseWindow(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? top_of_hwnd((vb_i32)A(&a[0])) : nullptr; if (w) w->showMinimized(); return w != nullptr; }
static vb_i64 u_FlashWindow(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? top_of_hwnd((vb_i32)A(&a[0])) : nullptr; if (w) QApplication::alert(w, 1000); return 0; }
static vb_i64 u_GetWindow(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc > 1 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w) return 0;
    int cmd = (int)A(&a[1]);
    if (cmd == 5) {                             /* GW_CHILD: the first control */
        for (QWidget *c : client_of(w)->findChildren<QWidget *>(QString(), Qt::FindDirectChildrenOnly)) {
            Ctl *k = ctl_of_widget(c);
            if (k && k->w == c) return hwnd_of(c);
        }
        return 0;
    }
    if (cmd == 4) return u_GetParent(1, a);     /* GW_OWNER */
    return 0;
}
static vb_i64 u_GetClassNameA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    QWidget *w = widget_of_hwnd((vb_i32)A(&a[0]));
    if (!w) return 0;
    QString n = "ThunderRT6FormDC";
    if (!dynamic_cast<FormWidget *>(w)) {
        Ctl *c = ctl_of_widget(w);
        n = c ? "ThunderRT6" + QString::fromStdString(c->type) : "Static";
        if (c && c->kind == K_TEXT) n = "ThunderRT6TextBox";
    }
    return put_astr(&a[1], A(&a[2]), n);
}
static vb_i64 u_WindowFromPoint(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 1) return 0;
    QPoint p = argc >= 2 ? QPoint((int)A(&a[0]), (int)A(&a[1])) : cursor_pos;
    for (QWidget *w = QApplication::widgetAt(p); w; w = w->parentWidget()) {
        if (dynamic_cast<FormWidget *>(w)) return hwnd_of(w);
        Ctl *c = ctl_of_widget(w);
        if (c && c->w == w) return hwnd_of(w);
    }
    return 0;
}
static vb_i64 u_GetFocus(vb_i32, vb_dllarg *)
{
    for (QWidget *w = QApplication::focusWidget(); w; w = w->parentWidget()) {
        Ctl *c = ctl_of_widget(w);
        if (c && c->w == w) return hwnd_of(w);
        if (dynamic_cast<FormWidget *>(w)) return hwnd_of(w);
    }
    return 0;
}
static vb_i64 u_SetFocus(vb_i32 argc, vb_dllarg *a)
{
    vb_i64 old = u_GetFocus(0, nullptr);
    QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (w) { w->activateWindow(); w->setFocus(); }
    return old;
}
static vb_i32 capture_hwnd;
static vb_i64 u_SetCapture(vb_i32 argc, vb_dllarg *a) { vb_i32 old = capture_hwnd; capture_hwnd = argc ? (vb_i32)A(&a[0]) : 0; return old; }
static vb_i64 u_GetCapture(vb_i32, vb_dllarg *) { return capture_hwnd; }
static vb_i64 u_ReleaseCapture(vb_i32, vb_dllarg *) { capture_hwnd = 0; return 1; }
static vb_i64 u_InvalidateRect(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr; if (w) w->update(); return 1; }
static vb_i64 u_UpdateWindow(vb_i32 argc, vb_dllarg *a) { QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr; if (w) w->repaint(); return 1; }
static QPointer<QWidget> locked_window;
static vb_i64 u_LockWindowUpdate(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (locked_window) { locked_window->setUpdatesEnabled(true); locked_window = nullptr; }
    if (w) { w->setUpdatesEnabled(false); locked_window = w; }
    return 1;
}
static vb_i64 u_EnumWindows(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 1) return 0;
    vb_i64 proc = A(&a[0]);
    vb_i32 lp = argc > 1 ? (vb_i32)A(&a[1]) : 0;
    if (!vb_callback_fn(proc)) return 0;
    std::vector<vb_i32> hs;
    for (vbobj f : loaded_forms) if (fstate(f)->w) hs.push_back(hwnd_of(fstate(f)->w));
    for (vb_i32 h : hs) {
        vb_i32 args[2] = { h, lp };
        if (!vb_callback(proc, 2, args) || vb_errpend) break;
    }
    return 1;
}
static vb_i64 u_EnumChildWindows(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    QWidget *w = widget_of_hwnd((vb_i32)A(&a[0]));
    vb_i64 proc = A(&a[1]);
    vb_i32 lp = argc > 2 ? (vb_i32)A(&a[2]) : 0;
    if (!w || !vb_callback_fn(proc)) return 0;
    std::vector<vb_i32> hs;
    for (QWidget *c : w->findChildren<QWidget *>()) { Ctl *k = ctl_of_widget(c); if (k && k->w == c) hs.push_back(hwnd_of(c)); }
    for (vb_i32 h : hs) {
        vb_i32 args[2] = { h, lp };
        if (!vb_callback(proc, 2, args) || vb_errpend) break;
    }
    return 1;
}

/* window styles: GetWindowLong / SetWindowLong; WS_EX_LAYERED + SetLayeredWindowAttributes */
static std::map<vb_i32, std::map<int, vb_i32>> wlongs;
static vb_i32 default_wlong(QWidget *w, int idx)
{
    bool top = w && w->isWindow();
    switch (idx) {
    case -16: return top ? (vb_i32)0x16CF0000 : (vb_i32)0x56010000;     /* GWL_STYLE */
    case -20: return top ? 0x00040100 : 0;                               /* GWL_EXSTYLE */
    case -4: return (vb_i32)(0x7FFD0000 | (hwnd_of(w) & 0xFFFF));          /* GWL_WNDPROC: a number of its own */
    case -6: return 0x00400000;                                          /* GWL_HINSTANCE */
    case -8: case -12: return 0;
    }
    return 0;
}
static vb_i64 u_GetWindowLongA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    vb_i32 h = (vb_i32)A(&a[0]);
    int idx = (int)A(&a[1]);
    QWidget *w = widget_of_hwnd(h);
    if (!w) return 0;
    auto it = wlongs[h].find(idx);
    return it != wlongs[h].end() ? it->second : default_wlong(w, idx);
}
static vb_i64 u_SetWindowLongA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    vb_i32 h = (vb_i32)A(&a[0]), v = (vb_i32)A(&a[2]);
    int idx = (int)A(&a[1]);
    QWidget *w = widget_of_hwnd(h);
    if (!w) return 0;
    vb_i32 old = (vb_i32)u_GetWindowLongA(2, a);
    wlongs[h][idx] = v;
    QWidget *t = w->window();
    if (idx == -16 && w->isWindow()) {
        bool caption = (v & 0x00C00000) != 0, was = (old & 0x00C00000) != 0;
        if (caption != was) {
            bool vis = t->isVisible();
            t->setWindowFlag(Qt::FramelessWindowHint, !caption);
            if (vis) t->show();
        }
    }
    if (idx == -20 && w->isWindow() && ((v ^ old) & 0x8)) {   /* WS_EX_TOPMOST */
        bool vis = t->isVisible();
        t->setWindowFlag(Qt::WindowStaysOnTopHint, v & 0x8);
        if (vis) t->show();
    }
    return old;
}
static vb_i64 u_SetLayeredWindowAttributes(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    QWidget *w = top_of_hwnd((vb_i32)A(&a[0]));
    if (!w) return 0;
    long key = (long)A(&a[1]);
    int alpha = (int)A(&a[2]) & 0xFF, flags = (int)A(&a[3]);
    if (flags & 2) w->setWindowOpacity(alpha / 255.0);
    if (flags & 1) {                            /* LWA_COLORKEY: the pixels of that color are holes */
        auto *fw = dynamic_cast<FormWidget *>(w);
        if (fw) {
            FormState *st = fstate(fw->form);
            QImage img = w->grab().toImage();
            if (!img.isNull()) {
                QBitmap m = QBitmap::fromImage(img.createMaskFromColor(qcolor(key & 0xFFFFFF).rgb(), Qt::MaskOutColor));
                w->setMask(m);
            }
            (void)st;
        }
    }
    return 1;
}

/* the system menu: a handle; removing its items changes nothing on Linux */
static vb_i64 u_GetSystemMenu(vb_i32 argc, vb_dllarg *a) { return argc ? (0x00A00000 | ((vb_i32)A(&a[0]) & 0xFFFF)) : 0; }
static vb_i64 u_GetMenuItemID(vb_i32, vb_dllarg *) { return -1; }
static vb_i64 u_ShowCursor(vb_i32 argc, vb_dllarg *a)
{
    static int count = 0;
    bool show = argc && A(&a[0]);
    int before = count;
    count += show ? 1 : -1;
    if (before >= 0 && count < 0) QApplication::setOverrideCursor(Qt::BlankCursor);
    if (before < 0 && count >= 0) QApplication::restoreOverrideCursor();
    return count;
}
static vb_i64 u_SystemParametersInfoA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    int action = (int)A(&a[0]);
    if (action == 0x30) {                       /* SPI_GETWORKAREA: a RECT */
        vb_i32 *r = (vb_i32 *)vb_dllarg_ptr(&a[2], 16);
        QScreen *s = QGuiApplication::primaryScreen();
        QRect g = s ? s->availableGeometry() : QRect(0, 0, 1024, 738);
        if (r) { r[0] = g.left(); r[1] = g.top(); r[2] = g.right() + 1; r[3] = g.bottom() + 1; }
        return 1;
    }
    if (action == 0x10 || action == 0x72 || action == 0x26) {   /* SPI_GETSCREENSAVEACTIVE, ... a BOOL */
        vb_i32 *p = (vb_i32 *)vb_dllarg_ptr(&a[2], 4);
        if (p) *p = 0;
    }
    return 1;                                   /* SPI_SETDESKWALLPAPER and the rest: done */
}
static long sys_color(int i)
{
    if (dark_theme && i >= 0 && i < (int)(sizeof sys_dark / sizeof sys_dark[0])) {   /* the dark theme */
        unsigned rgb = sys_dark[i];
        return (long)(((rgb >> 16) & 255) | (rgb & 0xFF00) | ((rgb & 255) << 16));
    }
    QPalette p = QApplication::palette();
    auto vbc = [](const QColor &c) { return (long)(c.red() | (c.green() << 8) | (c.blue() << 16)); };
    switch (i) {
    case 1: return 0x6B5B3A;                    /* COLOR_BACKGROUND */
    case 2: return 0xD1B499; case 3: return 0xDBCDBF;   /* active / inactive caption */
    case 4: return vbc(p.color(QPalette::Window));      /* COLOR_MENU */
    case 5: return vbc(p.color(QPalette::Base));        /* COLOR_WINDOW */
    case 6: return 0x646464;
    case 7: case 8: return vbc(p.color(QPalette::Text));
    case 9: return 0;
    case 10: case 11: return 0xB4B4B4;
    case 12: return 0xABABAB;
    case 13: return vbc(p.color(QPalette::Highlight));
    case 14: return vbc(p.color(QPalette::HighlightedText));
    case 15: return vbc(p.color(QPalette::Button));     /* COLOR_BTNFACE */
    case 16: return vbc(p.color(QPalette::Mid));
    case 17: return 0x6D6D6D;
    case 18: return vbc(p.color(QPalette::ButtonText));
    case 19: return 0x544E43;
    case 20: return 0xFFFFFF;
    case 21: return vbc(p.color(QPalette::Shadow));
    case 22: return vbc(p.color(QPalette::Light));
    case 23: return 0;
    case 24: return 0xE1FFFF;                   /* COLOR_INFOBK */
    case 26: return 0xCC6600;                   /* COLOR_HOTLIGHT */
    case 27: return 0xEAD1B9; case 28: return 0xF2E4D7;
    }
    return 0;
}
static vb_i64 u_GetSysColor(vb_i32 argc, vb_dllarg *a) { return argc ? sys_color((int)A(&a[0]) & 0xFF) : 0; }
static vb_i64 o_OleTranslateColor(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0x80070057;
    vb_i32 c = (vb_i32)A(&a[0]);
    vb_i32 *out = (vb_i32 *)vb_dllarg_ptr(&a[2], 4);
    long v = (c & 0x80000000) ? sys_color(c & 0xFF) : (long)(c & 0xFFFFFF);
    if (out) *out = (vb_i32)v;
    return 0;
}

/* ---------------------------------------------------------------- SendMessage to the controls */
static QAbstractScrollArea *scroll_of(QWidget *w) { return qobject_cast<QAbstractScrollArea *>(w); }
static vb_i64 send_message(vb_i32 h, vb_u32 msg, vb_i64 wp, vb_dllarg *lpa)
{
    QWidget *w = widget_of_hwnd(h);
    if (!w) return 0;
    auto lp_int = [&]() -> vb_i64 { return lpa ? A(lpa) : 0; };
    auto lp_text = [&]() -> QString { return lpa ? astr(lpa) : QString(); };
    QWidget *top = w->window();
    switch (msg) {
    case 0xA1:                                  /* WM_NCLBUTTONDOWN, HTCAPTION: the user drags the window */
        if (wp == 2 && top->windowHandle()) {
#if QT_VERSION >= QT_VERSION_CHECK(5, 15, 0)
            top->windowHandle()->startSystemMove();
#endif
        }
        return 0;
    case 0x112:                                 /* WM_SYSCOMMAND */
        switch (wp & 0xFFF0) {
        case 0xF010:
#if QT_VERSION >= QT_VERSION_CHECK(5, 15, 0)
            if (top->windowHandle()) top->windowHandle()->startSystemMove();
#endif
            return 0;
        case 0xF060: top->close(); return 0;
        case 0xF020: top->showMinimized(); return 0;
        case 0xF030: top->showMaximized(); return 0;
        case 0xF120: top->showNormal(); return 0;
        }
        return 0;
    case 0x10: top->close(); return 0;          /* WM_CLOSE */
    case 0x0B: w->setUpdatesEnabled(wp != 0); if (wp) w->update(); return 0;   /* WM_SETREDRAW */
    case 0x0C: {                                /* WM_SETTEXT */
        vb_dllarg args[2] = { { VBD_INT, h, 0, 0, 0 }, lpa ? *lpa : vb_dllarg{ VBD_INT, 0, 0, 0, 0 } };
        return u_SetWindowTextA(2, args);
    }
    case 0x0D: return lpa ? put_astr(lpa, wp, window_text(w)) : 0;   /* WM_GETTEXT */
    case 0x0E: return window_text(w).size();   /* WM_GETTEXTLENGTH */
    }
    /* text boxes */
    auto *le = qobject_cast<QLineEdit *>(w);
    auto *pe = qobject_cast<QPlainTextEdit *>(w);
    auto *te = qobject_cast<QTextEdit *>(w);
    if (le || pe || te) {
        QString text = le ? le->text() : pe ? pe->toPlainText() : te->toPlainText();
        auto cursor = [&]() { return pe ? pe->textCursor() : te->textCursor(); };
        auto line_start = [&](int line) { int p = 0; for (int i = 0; i < line; i++) { int n = text.indexOf('\n', p); if (n < 0) return -1; p = n + 1; } return p; };
        switch (msg) {
        case 0x300: if (le) le->cut(); else if (pe) pe->cut(); else te->cut(); return 0;
        case 0x301: if (le) le->copy(); else if (pe) pe->copy(); else te->copy(); return 0;
        case 0x302: {                           /* WM_PASTE: the text of the clipboard (no clipboard owner: nothing) */
            QString t = QGuiApplication::clipboard()->text();
            if (le) le->insert(t); else if (pe) pe->insertPlainText(t); else te->insertPlainText(t);
            return 0;
        }
        case 0x303: if (le) le->del(); else { QTextCursor c = cursor(); c.removeSelectedText(); } return 0;
        case 0x304: case 0xC7: if (le) le->undo(); else if (pe) pe->undo(); else te->undo(); return 1;
        case 0xC6: return le ? le->isUndoAvailable() : pe ? pe->document()->isUndoAvailable() : te->document()->isUndoAvailable();
        case 0xCD: if (pe) pe->document()->clearUndoRedoStacks(); else if (te) te->document()->clearUndoRedoStacks(); return 0;
        case 0xB1: {                            /* EM_SETSEL */
            int s = (int)wp, e = (int)lp_int();
            if (s < 0) { if (le) le->deselect(); return 0; }
            if (e < 0 || e > text.size()) e = text.size();
            if (le) le->setSelection(s, e - s);
            else { QTextCursor c = cursor(); c.setPosition(s); c.setPosition(e, QTextCursor::KeepAnchor); if (pe) pe->setTextCursor(c); else te->setTextCursor(c); }
            return 0;
        }
        case 0xB0: {                            /* EM_GETSEL */
            int s = le ? (le->hasSelectedText() ? le->selectionStart() : le->cursorPosition()) : cursor().selectionStart();
            int e = le ? s + le->selectedText().size() : cursor().selectionEnd();
            return (s & 0xFFFF) | ((vb_i64)(e & 0xFFFF) << 16);
        }
        case 0xB7: if (pe) pe->ensureCursorVisible(); else if (te) te->ensureCursorVisible(); return 0;   /* EM_SCROLLCARET */
        case 0xBA: return le ? 1 : text.count('\n') + 1;    /* EM_GETLINECOUNT */
        case 0xBB: {                            /* EM_LINEINDEX */
            int line = (int)wp;
            if (line < 0) line = text.left(le ? le->cursorPosition() : cursor().position()).count('\n');
            return line_start(line);
        }
        case 0xC1: {                            /* EM_LINELENGTH (wp: a character of the line) */
            int p = (int)wp < 0 ? 0 : (int)wp, s = text.lastIndexOf('\n', p - 1) + 1, e = text.indexOf('\n', s);
            return (e < 0 ? text.size() : e) - s;
        }
        case 0xC9: {                            /* EM_LINEFROMCHAR */
            int p = (int)wp < 0 ? (le ? le->cursorPosition() : cursor().position()) : (int)wp;
            return text.left(p).count('\n');
        }
        case 0xC4: {                            /* EM_GETLINE: lParam a buffer with its size in the first word */
            int s = line_start((int)wp);
            if (s < 0 || !lpa) return 0;
            int e = text.indexOf('\n', s);
            QString line = text.mid(s, (e < 0 ? text.size() : e) - s);
            if (line.endsWith('\r')) line.chop(1);
            char *b = vb_dllarg_str(lpa);
            int room = b ? (int)((unsigned char)b[0] | ((unsigned char)b[1] << 8)) : 0;
            if (room <= 0) room = 1024;
            vb_i32 mk = vb_mark(); int an = 0; char *t = vb_stoansi(vs(line), &an);
            if (an > room) an = room;
            if (b) memcpy(b, t, (size_t)an);
            vb_release(mk);
            return an;
        }
        case 0xB6: {                            /* EM_LINESCROLL */
            QAbstractScrollArea *sa = scroll_of(w);
            if (sa) sa->verticalScrollBar()->setValue(sa->verticalScrollBar()->value() + (int)lp_int());
            return 1;
        }
        case 0xCE: { QAbstractScrollArea *sa = scroll_of(w); return sa ? sa->verticalScrollBar()->value() : 0; }   /* EM_GETFIRSTVISIBLELINE */
        case 0xCF: if (le) le->setReadOnly(wp); else if (pe) pe->setReadOnly(wp); else te->setReadOnly(wp); return 1;
        case 0xC5: if (le) le->setMaxLength(wp > 0 ? (int)wp : 32767); return 0;    /* EM_LIMITTEXT */
        case 0xCC: if (le) le->setEchoMode(wp ? QLineEdit::Password : QLineEdit::Normal); return 0;
        case 0xB8: return le ? le->isModified() : pe ? pe->document()->isModified() : te->document()->isModified();
        case 0xB9: if (le) le->setModified(wp); else if (pe) pe->document()->setModified(wp); else te->document()->setModified(wp); return 0;
        case 0xC2: {                            /* EM_REPLACESEL */
            QString t = lp_text();
            if (le) le->insert(t); else { QTextCursor c = cursor(); c.insertText(t); }
            return 0;
        }
        case 0x115: case 0x114: {               /* WM_VSCROLL / WM_HSCROLL */
            QAbstractScrollArea *sa = scroll_of(w);
            if (!sa) return 0;
            QScrollBar *sb = msg == 0x115 ? sa->verticalScrollBar() : sa->horizontalScrollBar();
            switch (wp & 0xFFFF) {
            case 0: sb->triggerAction(QAbstractSlider::SliderSingleStepSub); break;
            case 1: sb->triggerAction(QAbstractSlider::SliderSingleStepAdd); break;
            case 2: sb->triggerAction(QAbstractSlider::SliderPageStepSub); break;
            case 3: sb->triggerAction(QAbstractSlider::SliderPageStepAdd); break;
            case 6: sb->setValue(sb->minimum()); break;
            case 7: sb->setValue(sb->maximum()); break;
            }
            return 0;
        }
        }
        return 0;
    }
    /* list boxes */
    if (auto *lw = qobject_cast<QListWidget *>(w)) {
        auto find = [&](bool exact) -> vb_i64 {
            QString t = lp_text();
            int start = (int)wp, n = lw->count();
            for (int k = 1; k <= n; k++) {
                int i = (start + k) % n;
                if (start < 0) i = k - 1;
                QString it = lw->item(i)->text();
                if (exact ? it.compare(t, Qt::CaseInsensitive) == 0 : it.startsWith(t, Qt::CaseInsensitive)) return i;
            }
            return -1;
        };
        switch (msg) {
        case 0x18B: return lw->count();         /* LB_GETCOUNT */
        case 0x188: return lw->currentRow();    /* LB_GETCURSEL */
        case 0x186: lw->setCurrentRow((int)wp); return wp < lw->count() ? wp : -1;   /* LB_SETCURSEL */
        case 0x18F: return find(false);         /* LB_FINDSTRING */
        case 0x1A2: return find(true);          /* LB_FINDSTRINGEXACT */
        case 0x18C: { vb_i64 i = find(false); if (i >= 0) lw->setCurrentRow((int)i); return i; }   /* LB_SELECTSTRING */
        case 0x18E: { QListWidgetItem *it = lw->itemAt(0, 0); return it ? lw->row(it) : 0; }      /* LB_GETTOPINDEX */
        case 0x197: if (wp >= 0 && wp < lw->count()) lw->scrollToItem(lw->item((int)wp), QAbstractItemView::PositionAtTop); return 0;
        case 0x189: return wp >= 0 && wp < lw->count() && lpa ? put_astr(lpa, 1 << 20, lw->item((int)wp)->text()) : -1;
        case 0x18A: return wp >= 0 && wp < lw->count() ? lw->item((int)wp)->text().size() : -1;
        case 0x187: return wp >= 0 && wp < lw->count() && lw->item((int)wp)->isSelected();   /* LB_GETSEL */
        case 0x190: return lw->selectedItems().size();                                       /* LB_GETSELCOUNT */
        case 0x185: {                           /* LB_SETSEL */
            int i = (int)lp_int();
            for (int k = 0; k < lw->count(); k++) if (i == -1 || k == i) lw->item(k)->setSelected(wp != 0);
            return 0;
        }
        case 0x1A1: return lw->count() ? lw->visualItemRect(lw->item(0)).height() : 16;      /* LB_GETITEMHEIGHT */
        case 0x194: lw->setHorizontalScrollBarPolicy(wp > lw->width() ? Qt::ScrollBarAlwaysOn : Qt::ScrollBarAsNeeded); return 0;
        case 0x115: if ((wp & 0xFFFF) == 7) lw->scrollToBottom(); else if ((wp & 0xFFFF) == 6) lw->scrollToTop(); return 0;
        }
        return 0;
    }
    /* combo boxes */
    if (auto *cb = qobject_cast<QComboBox *>(w)) {
        auto find = [&](bool exact) -> vb_i64 {
            QString t = lp_text();
            for (int i = (int)wp + 1 < 0 ? 0 : (int)wp + 1, k = 0; k < cb->count(); k++, i = (i + 1) % std::max(1, cb->count())) {
                QString it = cb->itemText(i);
                if (exact ? it.compare(t, Qt::CaseInsensitive) == 0 : it.startsWith(t, Qt::CaseInsensitive)) return i;
            }
            return -1;
        };
        switch (msg) {
        case 0x146: return cb->count();          /* CB_GETCOUNT */
        case 0x147: return cb->currentIndex();   /* CB_GETCURSEL */
        case 0x14E: cb->setCurrentIndex((int)wp); return wp < cb->count() ? wp : -1;
        case 0x14C: return find(false);          /* CB_FINDSTRING */
        case 0x158: return find(true);           /* CB_FINDSTRINGEXACT */
        case 0x14D: { vb_i64 i = find(false); if (i >= 0) cb->setCurrentIndex((int)i); return i; }
        case 0x14F: if (wp) cb->showPopup(); else cb->hidePopup(); return 1;   /* CB_SHOWDROPDOWN */
        case 0x157: return cb->view() && cb->view()->isVisible();            /* CB_GETDROPPEDSTATE */
        case 0x160: if (cb->view()) cb->view()->setMinimumWidth((int)wp); return wp;   /* CB_SETDROPPEDWIDTH */
        case 0x148: return wp >= 0 && wp < cb->count() && lpa ? put_astr(lpa, 1 << 20, cb->itemText((int)wp)) : -1;
        case 0x149: return wp >= 0 && wp < cb->count() ? cb->itemText((int)wp).size() : -1;
        case 0x141: if (cb->lineEdit()) cb->lineEdit()->setMaxLength(wp > 0 ? (int)wp : 32767); return 1;
        }
        return 0;
    }
    /* buttons */
    if (auto *b = qobject_cast<QAbstractButton *>(w)) {
        switch (msg) {
        case 0xF5: b->click(); return 0;        /* BM_CLICK */
        case 0xF0: return b->isChecked() ? 1 : 0;   /* BM_GETCHECK */
        case 0xF1: b->setChecked(wp != 0); return 0;
        }
        return 0;
    }
    /* the ListView (a QTreeWidget): column widths, grid lines, full-row select */
    if (auto *tv = qobject_cast<QTreeWidget *>(w)) {
        switch (msg) {
        case 0x101E:                            /* LVM_SETCOLUMNWIDTH */
            if (lp_int() < 0) tv->resizeColumnToContents((int)wp); else tv->setColumnWidth((int)wp, (int)lp_int());
            return 1;
        case 0x1004: return tv->topLevelItemCount();   /* LVM_GETITEMCOUNT */
        case 0x1036: case 0x1037: return 0;     /* LVM_SET/GETEXTENDEDLISTVIEWSTYLE */
        }
        return 0;
    }
    /* clicks sent to a control */
    if (msg == 0x201 || msg == 0x202) {
        QPoint at(lp_int() & 0xFFFF, (lp_int() >> 16) & 0xFFFF);
        QMouseEvent ev(msg == 0x201 ? QEvent::MouseButtonPress : QEvent::MouseButtonRelease, at, w->mapToGlobal(at),
                       Qt::LeftButton, msg == 0x201 ? Qt::LeftButton : Qt::NoButton, Qt::NoModifier);
        QCoreApplication::sendEvent(w, &ev);
    }
    return 0;
}
static vb_i64 u_SendMessageA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    return send_message((vb_i32)A(&a[0]), (vb_u32)A(&a[1]), argc > 2 ? A(&a[2]) : 0, argc > 3 ? &a[3] : nullptr);
}
static vb_i64 u_PostMessageA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    vb_i32 h = (vb_i32)A(&a[0]); vb_u32 msg = (vb_u32)A(&a[1]);
    vb_i64 wp = argc > 2 ? A(&a[2]) : 0, lp = argc > 3 ? A(&a[3]) : 0;
    QTimer::singleShot(0, [h, msg, wp, lp]() { vb_dllarg l = { VBD_INT, lp, 0, 0, 0 }; send_message(h, msg, wp, &l); });
    return 1;
}

/* ---------------------------------------------------------------- the tray: Shell_NotifyIcon
 * NOTIFYICONDATA (VB6 layout): cbSize, hWnd, uID, uFlags, uCallbackMessage, hIcon, szTip (64 or 128).
 * A click on the icon comes back as VB6 programs expect it: a MouseMove of the form whose X, in
 * pixels, is the message (WM_LBUTTONDOWN 0x201 ... WM_RBUTTONUP 0x205, WM_LBUTTONDBLCLK 0x203). */
static std::map<std::pair<vb_i32, vb_i32>, QSystemTrayIcon *> trays;
static void tray_message(vb_i32 hwnd, vb_i32 msg)
{
    QWidget *w = widget_of_hwnd(hwnd);
    vbobj o = obj_of_hwnd(hwnd);
    if (!w || !o) return;
    if (auto *fw = dynamic_cast<FormWidget *>(w)) {
        FormState *st = fstate(fw->form);
        vbvar av[4] = { vb_v_i(0), vb_v_i(0), vb_v_f((float)un_w(&st->g, msg)), vb_v_f(0) };
        fire_form(fw->form, "mousemove", 4, av);
    } else if (Ctl *c = ctl_of_widget(w)) {
        vbvar av[4] = { vb_v_i(0), vb_v_i(0), vb_v_f((float)(msg * TPP)), vb_v_f(0) };
        fire_ctl(c, "mousemove", 4, av);
    }
    if (vb_errpend) vb_unhandled();
}
static vb_i64 s_Shell_NotifyIconA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    int op = (int)A(&a[0]);
    unsigned char *d = (unsigned char *)vb_dllarg_ptr(&a[1], 24);
    if (!d) return 0;
    vb_i32 hwnd, id, flags, cbmsg, hicon;
    memcpy(&hwnd, d + 4, 4); memcpy(&id, d + 8, 4); memcpy(&flags, d + 12, 4); memcpy(&cbmsg, d + 16, 4); memcpy(&hicon, d + 20, 4);
    auto key = std::make_pair(hwnd, id);
    auto it = trays.find(key);
    if (op == 2) {                              /* NIM_DELETE */
        if (it != trays.end()) { it->second->hide(); it->second->deleteLater(); trays.erase(it); }
        return 1;
    }
    QSystemTrayIcon *t = it != trays.end() ? it->second : nullptr;
    if (!t) {
        if (op != 0) return 0;
        t = new QSystemTrayIcon(app);
        trays[key] = t;
        QWidget *w = top_of_hwnd(hwnd);
        if (w && !w->windowIcon().isNull()) t->setIcon(w->windowIcon());
        else t->setIcon(QApplication::windowIcon().isNull() ? QApplication::style()->standardIcon(QStyle::SP_ComputerIcon)
                                                              : QApplication::windowIcon());
        QObject::connect(t, &QSystemTrayIcon::activated, [hwnd, key](QSystemTrayIcon::ActivationReason r) {
            auto f = trays.find(key);
            if (f == trays.end()) return;
            vb_i32 cb = f->second->property("vbmsg").toInt();
            if (!cb) return;
            if (cb != 0x200) {                  /* another callback message: the procedure of a subclassed window gets it */
                return;
            }
            switch (r) {
            case QSystemTrayIcon::Trigger: tray_message(hwnd, 0x201); tray_message(hwnd, 0x202); break;
            case QSystemTrayIcon::DoubleClick: tray_message(hwnd, 0x203); break;
            case QSystemTrayIcon::Context: tray_message(hwnd, 0x204); tray_message(hwnd, 0x205); break;
            case QSystemTrayIcon::MiddleClick: tray_message(hwnd, 0x207); tray_message(hwnd, 0x208); break;
            default: break;
            }
        });
    }
    if (flags & 1) t->setProperty("vbmsg", cbmsg);
    if (flags & 4) {
        char tip[129]; memcpy(tip, d + 24, 128); tip[128] = 0;
        vb_i32 mk = vb_mark();
        t->setToolTip(qs(vb_sfromansi(tip, (int)strnlen(tip, 128))));
        vb_release(mk);
    }
    t->show();
    return 1;
}

/* ---------------------------------------------------------------- comdlg32: the file and color dialogs
 * OPENFILENAME (VB6 layout): 0 lStructSize, 12 lpstrFilter, 24 nFilterIndex, 28 lpstrFile, 32 nMaxFile,
 * 36 lpstrFileTitle, 40 nMaxFileTitle, 44 lpstrInitialDir, 48 lpstrTitle, 52 Flags, 56 nFileOffset,
 * 58 nFileExtension, 60 lpstrDefExt */
static const char *ptr_str(const unsigned char *d, int off)
{
    vb_i32 p; memcpy(&p, d + off, 4);
    return p ? (const char *)vb_ptr_resolve(p, 1) : nullptr;
}
static QString ansi_q(const char *s, int n = -1)
{
    if (!s) return QString();
    vb_i32 mk = vb_mark();
    QString r = qs(vb_sfromansi(s, n < 0 ? (int)strlen(s) : n));
    vb_release(mk);
    return r;
}
static vb_i64 file_dialog(vb_i32 argc, vb_dllarg *a, bool save)
{
    unsigned char *d = argc ? (unsigned char *)vb_dllarg_ptr(&a[0], 76) : nullptr;
    if (!d) return 0;
    if (monkey_on) { monkey_log(QString("GetOpenFileName: Cancel")); return 0; }
    QStringList filters;
    if (const char *f = ptr_str(d, 12)) {       /* "Text (*.txt)\0*.txt\0\0" */
        while (*f) {
            QString name = ansi_q(f); f += strlen(f) + 1;
            if (!*f) { filters << name; break; }
            QString pat = ansi_q(f); f += strlen(f) + 1;
            filters << name + " (" + pat.replace(';', ' ') + ")";
        }
    }
    vb_i32 idx; memcpy(&idx, d + 24, 4);
    QString sel = idx > 0 && idx <= filters.size() ? filters[idx - 1] : QString();
    vb_i32 flags; memcpy(&flags, d + 52, 4);
    QString dir = ansi_q(ptr_str(d, 44)).replace('\\', '/'), title = ansi_q(ptr_str(d, 48));
    const char *fbuf = ptr_str(d, 28);
    QString start = fbuf ? ansi_q(fbuf) : QString();
    if (!dir.isEmpty()) start = dir + (start.isEmpty() ? "" : "/" + start);
    QStringList files;
    if (save) {
        QString f = QFileDialog::getSaveFileName(dialog_parent(), title, start, filters.join(";;"), &sel);
        if (!f.isEmpty()) files << f;
    } else if (flags & 0x200)
        files = QFileDialog::getOpenFileNames(dialog_parent(), title, start, filters.join(";;"), &sel);
    else {
        QString f = QFileDialog::getOpenFileName(dialog_parent(), title, start, filters.join(";;"), &sel);
        if (!f.isEmpty()) files << f;
    }
    if (files.isEmpty()) return 0;
    if (save) {
        const char *ext = ptr_str(d, 60);
        if (ext && *ext && QFileInfo(files[0]).suffix().isEmpty()) files[0] += "." + ansi_q(ext);
    }
    QString result;
    if (files.size() > 1) {                     /* several: the folder, then the names, separated by 0 */
        result = QFileInfo(files[0]).absolutePath();
        for (const QString &f : files) result += QChar(0) + QFileInfo(f).fileName();
        result += QChar(0);
    } else result = files[0];
    vb_i32 maxf; memcpy(&maxf, d + 32, 4);
    vb_i32 fp; memcpy(&fp, d + 28, 4);
    char *out = fp ? (char *)vb_ptr_resolve(fp, maxf > 0 ? maxf : 1) : nullptr;
    vb_i32 mk = vb_mark();
    int an = 0; char *t = vb_stoansi(vs(result), &an);
    if (out && maxf > 0) { int k = an < maxf - 1 ? an : maxf - 1; memcpy(out, t, (size_t)k); out[k] = 0; }
    QString name = QFileInfo(files[0]).fileName();
    vb_u16 off = (vb_u16)(files.size() > 1 ? QFileInfo(files[0]).absolutePath().size() + 1 : result.size() - name.size());
    int dot = name.lastIndexOf('.');
    vb_u16 extoff = (vb_u16)(dot >= 0 && files.size() == 1 ? off + dot + 1 : 0);
    memcpy(d + 56, &off, 2); memcpy(d + 58, &extoff, 2);
    vb_i32 tp, maxt; memcpy(&tp, d + 36, 4); memcpy(&maxt, d + 40, 4);
    char *tout = tp ? (char *)vb_ptr_resolve(tp, maxt > 0 ? maxt : 1) : nullptr;
    if (tout && maxt > 0) { char *tt = vb_stoansi(vs(name), &an); int k = an < maxt - 1 ? an : maxt - 1; memcpy(tout, tt, (size_t)k); tout[k] = 0; }
    for (int i = 0; i < filters.size(); i++) if (filters[i] == sel) { vb_i32 n = i + 1; memcpy(d + 24, &n, 4); }
    vb_release(mk);
    return 1;
}
static vb_i64 c_GetOpenFileNameA(vb_i32 argc, vb_dllarg *a) { return file_dialog(argc, a, false); }
static vb_i64 c_GetSaveFileNameA(vb_i32 argc, vb_dllarg *a) { return file_dialog(argc, a, true); }
static vb_i64 c_ChooseColorA(vb_i32 argc, vb_dllarg *a)
{
    unsigned char *d = argc ? (unsigned char *)vb_dllarg_ptr(&a[0], 36) : nullptr;
    if (!d) return 0;
    if (monkey_on) { monkey_log("ChooseColor: Cancel"); return 0; }
    vb_i32 c; memcpy(&c, d + 12, 4);
    QColor q = QColorDialog::getColor(qcolor(c & 0xFFFFFF), dialog_parent());
    if (!q.isValid()) return 0;
    c = (vb_i32)vbcolor(q.rgb());
    memcpy(d + 12, &c, 4);
    return 1;
}
static vb_i64 c_CommDlgExtendedError(vb_i32, vb_dllarg *) { return 0; }
#ifndef _WIN32                                  /* on Windows: the real one (its PIDLs go on to SHGetPathFromIDList) */
/* SHBrowseForFolder: BROWSEINFO 0 hOwner, 4 pidlRoot, 8 pszDisplayName, 12 lpszTitle, 16 ulFlags */
static vb_i64 s_SHBrowseForFolderA(vb_i32 argc, vb_dllarg *a)
{
    unsigned char *d = argc ? (unsigned char *)vb_dllarg_ptr(&a[0], 32) : nullptr;
    if (!d) return 0;
    if (monkey_on) { monkey_log("SHBrowseForFolder: Cancel"); return 0; }
    QString title = ansi_q(ptr_str(d, 12));
    QString dir = QFileDialog::getExistingDirectory(dialog_parent(), title.isEmpty() ? QString("Browse for Folder") : title);
    if (dir.isEmpty()) return 0;
    vb_i32 dp; memcpy(&dp, d + 8, 4);
    char *disp = dp ? (char *)vb_ptr_resolve(dp, 1) : nullptr;
    vb_i32 mk = vb_mark();
    if (disp) { int an = 0; char *t = vb_stoansi(vs(QFileInfo(dir).fileName()), &an); memcpy(disp, t, (size_t)an + 1); }
    vb_release(mk);
    return vb_pidl_make(dir.toUtf8().constData());
}
#endif

/* ---------------------------------------------------------------- gdi32: objects, memory DCs, drawing
 * Pens, brushes, bitmaps, fonts and regions are handles of the runtime. A memory DC (CreateCompatibleDC)
 * draws on the bitmap selected into it; BitBlt between it and a form or PictureBox works as on Windows. */
struct GdiObj {
    int type = 0;                               /* 1 pen, 2 brush, 3 bitmap, 4 region, 5 font */
    QColor color; int style = 0, width = 1;
    Gfx *bmp = nullptr;                         /* a bitmap */
    QRegion rgn;
    QFont font;
    QImage pattern;                             /* a brush of CreatePatternBrush */
};
struct DcState {
    vb_i32 pen = 0, brush = 0, bitmap = 0, font = 0;
    QColor text = Qt::black, back = Qt::white; int bkmode = 2;   /* OPAQUE */
    QPoint pos; bool memory = false;
    Gfx *own = nullptr;                         /* the 1x1 bitmap of a new memory DC */
};
static std::map<vb_i32, GdiObj> gdi;
static std::map<vb_i32, DcState> dcs;
static vb_i32 gdi_next = 0x00300010;
static vb_i32 gdi_new(const GdiObj &o) { vb_i32 h = gdi_next; gdi_next += 4; gdi[h] = o; return h; }
static GdiObj *gdi_of(vb_i32 h) { auto it = gdi.find(h); return it == gdi.end() ? nullptr : &it->second; }
static Gfx *new_bitmap(int w, int h)
{
    Gfx *g = new Gfx;
    g->pw = std::max(1, w); g->ph = std::max(1, h);
    g->img = QImage(g->pw, g->ph, QImage::Format_RGB32);
    g->img.fill(Qt::black);
    g->scalemode = 3; gfx_rescale(g);
    return g;
}
static Gfx *memdc_gfx(vb_i32 h)
{
    auto it = dcs.find(h);
    if (it == dcs.end() || !it->second.memory) return nullptr;
    GdiObj *b = gdi_of(it->second.bitmap);
    return b && b->bmp ? b->bmp : it->second.own;
}
static DcState &dc_state(vb_i32 h) { return dcs[h]; }
/* a Picture's Handle (StdPicture) as a bitmap */
static vb_i32 bitmap_of_pic(vb_i32 h)
{
    auto it = pic_handles.find(h);
    if (it == pic_handles.end() || !it->second) return 0;
    for (auto &g : gdi) if (g.second.type == 3 && g.second.style == h) return g.first;
    GdiObj o; o.type = 3; o.style = h;
    o.bmp = new_bitmap(it->second->img.width(), it->second->img.height());
    o.bmp->img = it->second->img.convertToFormat(QImage::Format_RGB32);
    return gdi_new(o);
}
static vb_i64 g_CreateCompatibleDC(vb_i32, vb_dllarg *)
{
    static vb_i32 next = 0x00400010;
    vb_i32 h = next; next += 4;
    DcState &s = dcs[h];
    s.memory = true;
    s.own = new_bitmap(1, 1);
    return h;
}
static vb_i64 g_CreateDCA(vb_i32, vb_dllarg *) { return DESKTOP; }
static vb_i64 g_DeleteDC(vb_i32 argc, vb_dllarg *a)
{
    auto it = argc ? dcs.find((vb_i32)A(&a[0])) : dcs.end();
    if (it == dcs.end()) return 1;
    delete it->second.own;
    dcs.erase(it);
    return 1;
}
static vb_i64 g_CreateCompatibleBitmap(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    GdiObj o; o.type = 3; o.bmp = new_bitmap((int)A(&a[1]), (int)A(&a[2]));
    return gdi_new(o);
}
static vb_i64 g_CreateBitmap(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    GdiObj o; o.type = 3; o.bmp = new_bitmap((int)A(&a[0]), (int)A(&a[1]));
    return gdi_new(o);
}
static vb_i64 g_CreatePen(vb_i32 argc, vb_dllarg *a)
{
    GdiObj o; o.type = 1;
    o.style = argc > 0 ? (int)A(&a[0]) : 0; o.width = argc > 1 ? std::max(1, (int)A(&a[1])) : 1;
    o.color = qcolor(argc > 2 ? (long)A(&a[2]) & 0xFFFFFF : 0);
    return gdi_new(o);
}
static vb_i64 g_CreateSolidBrush(vb_i32 argc, vb_dllarg *a)
{
    GdiObj o; o.type = 2; o.color = qcolor(argc ? (long)A(&a[0]) & 0xFFFFFF : 0);
    return gdi_new(o);
}
static vb_i64 g_CreateHatchBrush(vb_i32 argc, vb_dllarg *a)
{
    GdiObj o; o.type = 2; o.style = 2; o.width = argc ? (int)A(&a[0]) : 0;
    o.color = qcolor(argc > 1 ? (long)A(&a[1]) & 0xFFFFFF : 0);
    return gdi_new(o);
}
static Gfx *bitmap_gfx(vb_i32 h);
static vb_i64 g_CreatePatternBrush(vb_i32 argc, vb_dllarg *a)   /* a brush that tiles a bitmap */
{
    Gfx *g = argc ? bitmap_gfx((vb_i32)A(&a[0])) : nullptr;
    if (!g) return 0;
    GdiObj o; o.type = 2; o.style = 3; o.pattern = g->img.copy();
    return gdi_new(o);
}
static vb_i64 g_GetStockObject(vb_i32 argc, vb_dllarg *a)
{
    int i = argc ? (int)A(&a[0]) : 0;
    static vb_i32 stock[20];
    if (i < 0 || i >= 20) return 0;
    if (!stock[i]) {
        GdiObj o;
        switch (i) {
        case 0: o.type = 2; o.color = Qt::white; break;
        case 1: o.type = 2; o.color = QColor(192, 192, 192); break;
        case 2: o.type = 2; o.color = Qt::gray; break;
        case 3: o.type = 2; o.color = Qt::darkGray; break;
        case 4: o.type = 2; o.color = Qt::black; break;
        case 5: o.type = 2; o.style = 1; break;           /* NULL_BRUSH */
        case 6: o.type = 1; o.color = Qt::white; break;
        case 7: o.type = 1; o.color = Qt::black; break;
        case 8: o.type = 1; o.style = 5; break;           /* NULL_PEN */
        default: o.type = 5; o.font = QApplication::font(); break;
        }
        stock[i] = gdi_new(o);
    }
    return stock[i];
}
static vb_i64 g_SelectObject(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]), h = (vb_i32)A(&a[1]);
    GdiObj *o = gdi_of(h);
    if (!o && pic_handles.count(h)) { h = bitmap_of_pic(h); o = gdi_of(h); }
    DcState &s = dc_state(dc);
    if (!o) return 0;
    vb_i32 old = 0;
    switch (o->type) {
    case 1: old = s.pen; s.pen = h; break;
    case 2: old = s.brush; s.brush = h; break;
    case 3: old = s.bitmap; s.bitmap = h; break;
    case 4: return 1;                           /* SIMPLEREGION */
    case 5: old = s.font; s.font = h; break;
    }
    return old ? old : 1;                       /* Windows gives back the old object (here: the first one is no object) */
}
static vb_i64 g_DeleteObject(vb_i32 argc, vb_dllarg *a)
{
    auto it = argc ? gdi.find((vb_i32)A(&a[0])) : gdi.end();
    if (it == gdi.end()) return 0;
    delete it->second.bmp;
    gdi.erase(it);
    return 1;
}
static vb_i64 g_GetObjectA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    vb_i32 h = (vb_i32)A(&a[0]);
    GdiObj *o = gdi_of(h);
    if (!o && pic_handles.count(h)) o = gdi_of(bitmap_of_pic(h));
    vb_i32 n = (vb_i32)A(&a[1]);
    unsigned char *p = (unsigned char *)vb_dllarg_ptr(&a[2], n > 0 ? n : 24);
    if (!o || !p || o->type != 3 || !o->bmp) return 0;
    memset(p, 0, (size_t)(n > 24 ? 24 : n));
    vb_i32 w = o->bmp->img.width(), ht = o->bmp->img.height(), wb = w * 4;
    vb_u16 planes = 1, bits = 32;
    memcpy(p + 4, &w, 4); memcpy(p + 8, &ht, 4); memcpy(p + 12, &wb, 4); memcpy(p + 16, &planes, 2); memcpy(p + 18, &bits, 2);
    return 24;
}
static vb_i64 g_SetTextColor(vb_i32 argc, vb_dllarg *a) { if (argc < 2) return 0; DcState &s = dc_state((vb_i32)A(&a[0])); long o = vbcolor(s.text.rgb()); s.text = qcolor((long)A(&a[1]) & 0xFFFFFF); return o; }
static vb_i64 g_SetBkColor(vb_i32 argc, vb_dllarg *a) { if (argc < 2) return 0; DcState &s = dc_state((vb_i32)A(&a[0])); long o = vbcolor(s.back.rgb()); s.back = qcolor((long)A(&a[1]) & 0xFFFFFF); return o; }
static vb_i64 g_SetBkMode(vb_i32 argc, vb_dllarg *a) { if (argc < 2) return 0; DcState &s = dc_state((vb_i32)A(&a[0])); int o = s.bkmode; s.bkmode = (int)A(&a[1]); return o; }
static vb_i64 g_SetStretchBltMode(vb_i32, vb_dllarg *) { return 1; }
/* a painter on the picture of a DC, with its pen and brush */
static bool dc_painter(vb_i32 dc, QPainter &p, Gfx **out)
{
    Gfx *g = gfx_of_dc(dc);
    if (!g || g->img.isNull()) return false;
    *out = g;
    p.begin(&g->img);
    DcState &s = dc_state(dc);
    GdiObj *pen = gdi_of(s.pen), *br = gdi_of(s.brush);
    if (pen && pen->style == 5) p.setPen(Qt::NoPen);
    else p.setPen(QPen(pen ? pen->color : QColor(Qt::black), pen ? pen->width : 1,
                       pen && pen->style == 1 ? Qt::DashLine : pen && pen->style == 2 ? Qt::DotLine : Qt::SolidLine));
    if (br && br->style == 1) p.setBrush(Qt::NoBrush);
    else if (br && br->style == 2) p.setBrush(QBrush(br->color, (Qt::BrushStyle)(Qt::HorPattern + std::min(5, br->width))));
    else if (br && br->style == 3) p.setBrush(QBrush(br->pattern));
    else p.setBrush(br ? QBrush(br->color) : QBrush(Qt::white));
    return true;
}
static void dc_done(QPainter &p, Gfx *g) { p.end(); gfx_update(g); }
static vb_i64 g_MoveToEx(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    DcState &s = dc_state((vb_i32)A(&a[0]));
    if (argc > 3) { vb_i32 *old = (vb_i32 *)vb_dllarg_ptr(&a[3], 8); if (old && a[3].kind == VBD_PTR) { old[0] = s.pos.x(); old[1] = s.pos.y(); } }
    s.pos = QPoint((int)A(&a[1]), (int)A(&a[2]));
    return 1;
}
static vb_i64 g_LineTo(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    QPoint to((int)A(&a[1]), (int)A(&a[2]));
    QPainter p; Gfx *g;
    if (dc_painter(dc, p, &g)) { p.drawLine(dc_state(dc).pos, to); dc_done(p, g); }
    dc_state(dc).pos = to;
    return 1;
}
static vb_i64 g_shape(vb_i32 argc, vb_dllarg *a, int kind)
{
    if (argc < 5) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    QRect r(QPoint((int)A(&a[1]), (int)A(&a[2])), QPoint((int)A(&a[3]) - 1, (int)A(&a[4]) - 1));
    QPainter p; Gfx *g;
    if (!dc_painter(dc, p, &g)) return 0;
    if (kind == 0) p.drawRect(r);
    else if (kind == 1) p.drawEllipse(r);
    else p.drawRoundedRect(r, argc > 5 ? A(&a[5]) / 2.0 : 4, argc > 6 ? A(&a[6]) / 2.0 : 4);
    dc_done(p, g);
    return 1;
}
static vb_i64 g_Rectangle(vb_i32 argc, vb_dllarg *a) { return g_shape(argc, a, 0); }
static vb_i64 g_Ellipse(vb_i32 argc, vb_dllarg *a) { return g_shape(argc, a, 1); }
static vb_i64 g_RoundRect(vb_i32 argc, vb_dllarg *a) { return g_shape(argc, a, 2); }
static QPolygon points_of(vb_dllarg *a, int n)
{
    QPolygon poly;
    vb_i32 *pt = n > 0 ? (vb_i32 *)vb_dllarg_ptr(a, n * 8) : nullptr;
    if (pt) for (int i = 0; i < n; i++) poly << QPoint(pt[2 * i], pt[2 * i + 1]);
    return poly;
}
static vb_i64 g_poly(vb_i32 argc, vb_dllarg *a, int kind)
{
    if (argc < 3) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    QPolygon poly = points_of(&a[1], (int)A(&a[2]));
    QPainter p; Gfx *g;
    if (poly.isEmpty() || !dc_painter(dc, p, &g)) return 0;
    if (kind == 0) p.drawPolygon(poly);
    else if (kind == 1) p.drawPolyline(poly);
    else {
        QPainterPath path(poly[0]);
        for (int i = 1; i + 2 < poly.size(); i += 3) path.cubicTo(poly[i], poly[i + 1], poly[i + 2]);
        p.setBrush(Qt::NoBrush);
        p.drawPath(path);
    }
    dc_done(p, g);
    return 1;
}
static vb_i64 g_Polygon(vb_i32 argc, vb_dllarg *a) { return g_poly(argc, a, 0); }
static vb_i64 g_Polyline(vb_i32 argc, vb_dllarg *a) { return g_poly(argc, a, 1); }
static vb_i64 g_PolyBezier(vb_i32 argc, vb_dllarg *a) { return g_poly(argc, a, 2); }
static vb_i64 g_FillRect(vb_i32 argc, vb_dllarg *a, bool frame)
{
    if (argc < 3) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    vb_i32 *r = (vb_i32 *)vb_dllarg_ptr(&a[1], 16);
    vb_i32 hb = (vb_i32)A(&a[2]);
    GdiObj *b = gdi_of(hb);
    QColor c = b ? b->color : (hb > 0 && hb <= 31 ? qcolor(sys_color(hb - 1)) : QColor(Qt::white));   /* COLOR_x + 1 */
    Gfx *g = gfx_of_dc(dc);
    if (!g || !r) return 0;
    QPainter p(&g->img);
    QRect rr(QPoint(r[0], r[1]), QPoint(r[2] - 1, r[3] - 1));
    if (frame) { p.setPen(c); p.setBrush(Qt::NoBrush); p.drawRect(rr.adjusted(0, 0, -1, -1)); }
    else if (b && b->style == 3 && !b->pattern.isNull()) p.fillRect(rr, QBrush(b->pattern));
    else p.fillRect(rr, c);
    p.end();
    gfx_update(g);
    return 1;
}
static vb_i64 u_FillRect(vb_i32 argc, vb_dllarg *a) { return g_FillRect(argc, a, false); }
static vb_i64 u_FrameRect(vb_i32 argc, vb_dllarg *a) { return g_FillRect(argc, a, true); }
static QFont dc_font(vb_i32 dc)
{
    GdiObj *f = gdi_of(dc_state(dc).font);
    if (f && f->type == 5) return f->font;
    Gfx *g = gfx_of_dc(dc);
    return g && g->font ? qfont(*g->font) : QApplication::font();
}
static vb_i64 g_TextOutA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    QString t = astr(&a[3]);
    if (argc > 4 && A(&a[4]) >= 0) t = t.left((int)A(&a[4]));
    Gfx *g = gfx_of_dc(dc);
    if (!g) return 0;
    DcState &s = dc_state(dc);
    QPainter p(&g->img);
    QFont f = dc_font(dc);
    p.setFont(f);
    QFontMetrics fm(f);
    QPoint at((int)A(&a[1]), (int)A(&a[2]));
    if (s.bkmode == 2) p.fillRect(QRect(at, QSize(fm.horizontalAdvance(t), fm.height())), s.back);
    p.setPen(s.text);
    p.drawText(at.x(), at.y() + fm.ascent(), t);
    p.end();
    gfx_update(g);
    return 1;
}
static vb_i64 u_DrawTextA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 5) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    QString t = astr(&a[1]);
    if (A(&a[2]) >= 0) t = t.left((int)A(&a[2]));
    vb_i32 *r = (vb_i32 *)vb_dllarg_ptr(&a[3], 16);
    int fl = (int)A(&a[4]);
    if (!r) return 0;
    QFont f = dc_font(dc);
    QFontMetrics fm(f);
    QRect rr(QPoint(r[0], r[1]), QPoint(r[2] - 1, r[3] - 1));
    int qf = (fl & 1 ? Qt::AlignHCenter : fl & 2 ? Qt::AlignRight : Qt::AlignLeft) |
             (fl & 4 ? Qt::AlignVCenter : fl & 8 ? Qt::AlignBottom : Qt::AlignTop) | (fl & 0x10 ? Qt::TextWordWrap : 0);
    if (!(fl & 0x20)) qf &= ~Qt::AlignVCenter, qf |= (fl & 4) ? Qt::AlignTop : 0;
    if (fl & 0x20) qf |= (fl & 4) ? Qt::AlignVCenter : 0;
    QRect br = fm.boundingRect(rr, qf, t);
    if (fl & 0x400) { r[2] = r[0] + br.width(); r[3] = r[1] + br.height(); return br.height(); }   /* DT_CALCRECT */
    Gfx *g = gfx_of_dc(dc);
    if (!g) return br.height();
    DcState &s = dc_state(dc);
    QPainter p(&g->img);
    p.setFont(f);
    if (s.bkmode == 2) p.fillRect(br, s.back);
    p.setPen(s.text);
    p.drawText(rr, qf, t);
    p.end();
    gfx_update(g);
    return br.height();
}
/* StdFont.hFont: a font object of the GDI emulation, one per font (it follows the font's changes) */
static vb_i32 font_handle(const VFont &f)
{
    static std::map<const VFont *, vb_i32> handles;
    GdiObj o; o.type = 5;
    o.font = qfont(f);
    auto it = handles.find(&f);
    if (it != handles.end() && gdi.count(it->second)) { gdi[it->second] = o; return it->second; }
    return handles[&f] = gdi_new(o);
}
static vb_i64 g_CreateFontA(vb_i32 argc, vb_dllarg *a)
{
    GdiObj o; o.type = 5;
    o.font = QApplication::font();
    if (argc >= 14) {
        int h = (int)A(&a[0]);
        if (h) o.font.setPixelSize(std::abs(h));
        int wt = (int)A(&a[4]);
        o.font.setBold(wt >= 600);
        o.font.setItalic(A(&a[5]) != 0); o.font.setUnderline(A(&a[6]) != 0); o.font.setStrikeOut(A(&a[7]) != 0);
        QString face = astr(&a[13]);
        if (!face.isEmpty()) o.font.setFamily(face);
    }
    return gdi_new(o);
}
/* msimg32 */
static vb_i64 m_TransparentBlt(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 11) return 0;
    Gfx *d = gfx_of_dc((vb_i32)A(&a[0]));
    vb_i32 hs = (vb_i32)A(&a[5]);
    int ox, oy, sw = (int)A(&a[8]), sh = (int)A(&a[9]);
    QImage src = source_image(hs, (int)A(&a[6]), (int)A(&a[7]), sw, sh, &ox, &oy);
    if (!d || src.isNull() || sw <= 0 || sh <= 0) return 0;
    QImage part = src.copy(ox, oy, sw, sh).convertToFormat(QImage::Format_ARGB32);
    QRgb key = qcolor((long)A(&a[10]) & 0xFFFFFF).rgb() & 0xFFFFFF;
    for (int y = 0; y < part.height(); y++) {
        QRgb *l = (QRgb *)part.scanLine(y);
        for (int x = 0; x < part.width(); x++) if ((l[x] & 0xFFFFFF) == key) l[x] = 0;
    }
    QPainter p(&d->img);
    p.drawImage(QRect((int)A(&a[1]), (int)A(&a[2]), (int)A(&a[3]), (int)A(&a[4])), part);
    p.end();
    gfx_update(d);
    return 1;
}
static vb_i64 m_AlphaBlend(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 11) return 0;
    Gfx *d = gfx_of_dc((vb_i32)A(&a[0]));
    int ox, oy, sw = (int)A(&a[8]), sh = (int)A(&a[9]);
    QImage src = source_image((vb_i32)A(&a[5]), (int)A(&a[6]), (int)A(&a[7]), sw, sh, &ox, &oy);
    if (!d || src.isNull()) return 0;
    vb_u32 bf = (vb_u32)A(&a[10]);
    double alpha = ((bf >> 16) & 0xFF) / 255.0;
    QPainter p(&d->img);
    p.setOpacity(alpha);
    p.drawImage(QRect((int)A(&a[1]), (int)A(&a[2]), (int)A(&a[3]), (int)A(&a[4])), src.copy(ox, oy, sw, sh));
    p.end();
    gfx_update(d);
    return 1;
}

/* regions: CreateRectRgn ..., CombineRgn, SetWindowRgn (the shape of a form or a control) */
static vb_i64 rgn_new(const QRegion &r) { GdiObj o; o.type = 4; o.rgn = r; return gdi_new(o); }
static vb_i64 g_CreateRectRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    return rgn_new(QRegion(QRect(QPoint((int)A(&a[0]), (int)A(&a[1])), QPoint((int)A(&a[2]) - 1, (int)A(&a[3]) - 1))));
}
static vb_i64 g_CreateEllipticRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    return rgn_new(QRegion(QRect(QPoint((int)A(&a[0]), (int)A(&a[1])), QPoint((int)A(&a[2]) - 1, (int)A(&a[3]) - 1)), QRegion::Ellipse));
}
static vb_i64 g_CreateRoundRectRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 6) return 0;
    QRect r(QPoint((int)A(&a[0]), (int)A(&a[1])), QPoint((int)A(&a[2]) - 1, (int)A(&a[3]) - 1));
    QPainterPath path;
    path.addRoundedRect(r, A(&a[4]) / 2.0, A(&a[5]) / 2.0);
    return rgn_new(QRegion(path.toFillPolygon().toPolygon()));
}
static vb_i64 g_CreatePolygonRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    return rgn_new(QRegion(points_of(&a[0], (int)A(&a[1])), argc > 2 && A(&a[2]) == 2 ? Qt::WindingFill : Qt::OddEvenFill));
}
static vb_i64 g_CreatePolyPolygonRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    int n = (int)A(&a[2]);
    vb_i32 *counts = n > 0 ? (vb_i32 *)vb_dllarg_ptr(&a[1], n * 4) : nullptr;
    if (!counts) return 0;
    int total = 0;
    for (int i = 0; i < n; i++) total += counts[i];
    QPolygon all = points_of(&a[0], total);
    QRegion r;
    int at = 0;
    for (int i = 0; i < n; i++) { r = r.united(QRegion(all.mid(at, counts[i]))); at += counts[i]; }
    return rgn_new(r);
}
static vb_i64 g_CombineRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    GdiObj *d = gdi_of((vb_i32)A(&a[0])), *x = gdi_of((vb_i32)A(&a[1])), *y = gdi_of((vb_i32)A(&a[2]));
    if (!d || !x) return 0;
    QRegion r1 = x->rgn, r2 = y ? y->rgn : QRegion();
    switch ((int)A(&a[3])) {
    case 1: d->rgn = r1.intersected(r2); break;
    case 2: d->rgn = r1.united(r2); break;
    case 3: d->rgn = r1.xored(r2); break;
    case 4: d->rgn = r1.subtracted(r2); break;
    default: d->rgn = r1; break;
    }
    return d->rgn.isEmpty() ? 1 : (d->rgn.rectCount() == 1 ? 2 : 3);   /* NULLREGION, SIMPLEREGION, COMPLEXREGION */
}
static vb_i64 g_OffsetRgn(vb_i32 argc, vb_dllarg *a)
{
    GdiObj *r = argc >= 3 ? gdi_of((vb_i32)A(&a[0])) : nullptr;
    if (!r) return 0;
    r->rgn.translate((int)A(&a[1]), (int)A(&a[2]));
    return r->rgn.rectCount() > 1 ? 3 : 2;
}
static vb_i64 g_PtInRegion(vb_i32 argc, vb_dllarg *a)
{
    GdiObj *r = argc >= 3 ? gdi_of((vb_i32)A(&a[0])) : nullptr;
    return r && r->rgn.contains(QPoint((int)A(&a[1]), (int)A(&a[2])));
}
static vb_i64 u_SetWindowRgn(vb_i32 argc, vb_dllarg *a)
{
    QWidget *w = argc >= 2 ? widget_of_hwnd((vb_i32)A(&a[0])) : nullptr;
    if (!w) return 0;
    GdiObj *r = gdi_of((vb_i32)A(&a[1]));
    if (!r) { w->clearMask(); return 1; }
    QRegion rg = r->rgn;
    if (w->isWindow() && !w->windowFlags().testFlag(Qt::FramelessWindowHint)) {
        /* a region of the whole window (with its frame): here the client area is the window */
        if (auto *fw = dynamic_cast<FormWidget *>(w)) {
            Frame f = frame_of(fstate(fw->form));
            rg.translate(-f.l, -(f.t));
            bool vis = w->isVisible();
            w->setWindowFlag(Qt::FramelessWindowHint, true);
            if (vis) w->show();
        }
    }
    w->setMask(rg);
    gdi.erase((vb_i32)A(&a[1]));               /* the window owns the region now */
    return 1;
}
static vb_i64 g_FillRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    Gfx *g = gfx_of_dc((vb_i32)A(&a[0]));
    GdiObj *r = gdi_of((vb_i32)A(&a[1])), *b = gdi_of((vb_i32)A(&a[2]));
    if (!g || !r) return 0;
    QPainter p(&g->img);
    p.setClipRegion(r->rgn);
    p.fillRect(r->rgn.boundingRect(), b ? b->color : QColor(Qt::black));
    p.end();
    gfx_update(g);
    return 1;
}
static vb_i64 u_LoadImageA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 6) return 0;
    if (!(A(&a[5]) & 0x10)) return 0;           /* LR_LOADFROMFILE only: there are no resources of Windows here */
    QString f = astr(&a[1]).replace('\\', '/');
    QImage img(f);
    if (img.isNull()) return 0;
    GdiObj o; o.type = 3;
    o.bmp = new_bitmap(img.width(), img.height());
    o.bmp->img = img.convertToFormat(QImage::Format_RGB32);
    return gdi_new(o);
}

/* ---------------------------------------------------------------- DIBs: the pixels of a bitmap in an array
 * BITMAPINFOHEADER: 0 biSize, 4 biWidth, 8 biHeight (> 0: the last row first), 12 biPlanes, 14 biBitCount,
 * 16 biCompression, 20 biSizeImage. 32 bits: B G R x; 24 bits: B G R, each row padded to 4 bytes. */
static Gfx *bitmap_gfx(vb_i32 h)
{
    GdiObj *o = gdi_of(h);
    if (!o && pic_handles.count(h)) o = gdi_of(bitmap_of_pic(h));
    return o && o->type == 3 ? o->bmp : nullptr;
}
static void dib_copy(QImage &img, unsigned char *bits, const unsigned char *bi, int start, int lines, bool to_bits)
{
    vb_i32 w, h; vb_u16 bpp;
    memcpy(&w, bi + 4, 4); memcpy(&h, bi + 8, 4); memcpy(&bpp, bi + 14, 2);
    bool bottomup = h > 0;
    int H = std::abs(h), stride = ((w * bpp + 31) / 32) * 4;
    for (int r = start; r < start + lines && r < H; r++) {
        int y = bottomup ? H - 1 - r : r;
        if (y < 0 || y >= img.height()) continue;
        unsigned char *row = bits + (size_t)r * stride;
        QRgb *line = (QRgb *)img.scanLine(y);
        for (int x = 0; x < w && x < img.width(); x++) {
            if (bpp == 32) {
                if (to_bits) { QRgb c = line[x]; row[4 * x] = qBlue(c); row[4 * x + 1] = qGreen(c); row[4 * x + 2] = qRed(c); row[4 * x + 3] = 0; }
                else line[x] = qRgb(row[4 * x + 2], row[4 * x + 1], row[4 * x]);
            } else if (bpp == 24) {
                if (to_bits) { QRgb c = line[x]; row[3 * x] = qBlue(c); row[3 * x + 1] = qGreen(c); row[3 * x + 2] = qRed(c); }
                else line[x] = qRgb(row[3 * x + 2], row[3 * x + 1], row[3 * x]);
            }
        }
    }
}
static vb_i64 g_GetDIBits(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 7) return 0;
    Gfx *g = bitmap_gfx((vb_i32)A(&a[1]));
    unsigned char *bi = (unsigned char *)vb_dllarg_ptr(&a[5], 40);
    if (!g || !bi) return 0;
    QImage &img = g->img;
    vb_i32 w, h; vb_u16 bpp;
    memcpy(&bpp, bi + 14, 2);
    unsigned char *bits = (unsigned char *)(a[4].kind == VBD_INT && !a[4].i ? nullptr : vb_dllarg_ptr(&a[4], 1));
    if (!bits || bpp == 0) {                    /* the size of the bitmap only */
        w = img.width(); h = img.height(); bpp = 32;
        vb_u16 planes = 1; vb_i32 size = w * h * 4, zero = 0;
        memcpy(bi + 4, &w, 4); memcpy(bi + 8, &h, 4); memcpy(bi + 12, &planes, 2); memcpy(bi + 14, &bpp, 2);
        memcpy(bi + 16, &zero, 4); memcpy(bi + 20, &size, 4);
        return img.height();
    }
    if (bpp != 24 && bpp != 32) return 0;
    int start = (int)A(&a[2]), lines = (int)A(&a[3]);
    dib_copy(img, bits, bi, start, lines, true);
    return lines;
}
static vb_i64 g_SetDIBits(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 7) return 0;
    Gfx *g = bitmap_gfx((vb_i32)A(&a[1]));
    unsigned char *bi = (unsigned char *)vb_dllarg_ptr(&a[5], 40), *bits = (unsigned char *)vb_dllarg_ptr(&a[4], 1);
    if (!g || !bi || !bits) return 0;
    int lines = (int)A(&a[3]);
    dib_copy(g->img, bits, bi, (int)A(&a[2]), lines, false);
    gfx_update(g);
    return lines;
}
static QImage dib_image(const unsigned char *bits, const unsigned char *bi)
{
    vb_i32 w, h;
    memcpy(&w, bi + 4, 4); memcpy(&h, bi + 8, 4);
    if (w <= 0 || h == 0 || w > 20000 || std::abs(h) > 20000) return QImage();
    QImage img(w, std::abs(h), QImage::Format_RGB32);
    dib_copy(img, (unsigned char *)bits, bi, 0, std::abs(h), false);
    return img;
}
static vb_i64 g_StretchDIBits(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 13) return 0;
    Gfx *d = gfx_of_dc((vb_i32)A(&a[0]));
    unsigned char *bits = (unsigned char *)vb_dllarg_ptr(&a[9], 1), *bi = (unsigned char *)vb_dllarg_ptr(&a[10], 40);
    if (!d || !bits || !bi) return 0;
    QImage src = dib_image(bits, bi);
    if (src.isNull()) return 0;
    int ws = (int)A(&a[7]), hs = (int)A(&a[8]), wd = (int)A(&a[3]), hd = (int)A(&a[4]);
    if (!ws || !hs || !wd || !hd) return 0;
    QImage part = src.copy((int)A(&a[5]), (int)A(&a[6]), std::abs(ws), std::abs(hs))
                     .mirrored((ws < 0) != (wd < 0), (hs < 0) != (hd < 0))
                     .scaled(std::abs(wd), std::abs(hd), Qt::IgnoreAspectRatio, Qt::FastTransformation);
    rop_blt(d, (int)A(&a[1]), (int)A(&a[2]), part, 0, 0, part.width(), part.height(), (vb_u32)A(&a[12]));
    return hs;
}
static vb_i64 g_SetDIBitsToDevice(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 12) return 0;
    Gfx *d = gfx_of_dc((vb_i32)A(&a[0]));
    unsigned char *bits = (unsigned char *)vb_dllarg_ptr(&a[9], 1), *bi = (unsigned char *)vb_dllarg_ptr(&a[10], 40);
    if (!d || !bits || !bi) return 0;
    QImage src = dib_image(bits, bi);
    if (src.isNull()) return 0;
    int w = (int)A(&a[3]), h = (int)A(&a[4]);
    rop_blt(d, (int)A(&a[1]), (int)A(&a[2]), src, (int)A(&a[5]), (int)A(&a[6]), w, h, 0xCC0020);
    return h;
}

/* ---------------------------------------------------------------- more drawing */
static vb_i64 g_CreateBrushIndirect(vb_i32 argc, vb_dllarg *a)
{
    vb_i32 *lb = argc ? (vb_i32 *)vb_dllarg_ptr(&a[0], 12) : nullptr;
    if (!lb) return 0;
    GdiObj o; o.type = 2;
    o.style = lb[0] == 1 ? 1 : lb[0] == 2 ? 2 : 0;   /* BS_SOLID, BS_NULL, BS_HATCHED */
    o.width = lb[2];
    o.color = qcolor(lb[1] & 0xFFFFFF);
    return gdi_new(o);
}
static vb_i64 g_CreateFontIndirectA(vb_i32 argc, vb_dllarg *a)
{
    unsigned char *lf = argc ? (unsigned char *)vb_dllarg_ptr(&a[0], 60) : nullptr;
    if (!lf) return 0;
    GdiObj o; o.type = 5;
    o.font = QApplication::font();
    vb_i32 h, wt; memcpy(&h, lf, 4); memcpy(&wt, lf + 16, 4);
    if (h) o.font.setPixelSize(std::abs(h));
    o.font.setBold(wt >= 600);
    o.font.setItalic(lf[20] != 0); o.font.setUnderline(lf[21] != 0); o.font.setStrikeOut(lf[22] != 0);
    char face[33]; memcpy(face, lf + 28, 32); face[32] = 0;
    if (face[0]) o.font.setFamily(ansi_q(face, (int)strnlen(face, 32)));
    return gdi_new(o);
}
static vb_i64 g_Pie(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 9) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    QRectF r(QPointF((double)A(&a[1]), (double)A(&a[2])), QPointF((double)A(&a[3]) - 1, (double)A(&a[4]) - 1));
    QPointF c = r.center();
    auto ang = [&](double x, double y) { return std::atan2(-(y - c.y()) * r.width() / std::max(1.0, r.height()), x - c.x()) * 180 / 3.14159265358979323846; };
    double a1 = ang((double)A(&a[5]), (double)A(&a[6])), a2 = ang((double)A(&a[7]), (double)A(&a[8]));
    double span = a2 - a1; if (span <= 0) span += 360;
    QPainter p; Gfx *g;
    if (!dc_painter(dc, p, &g)) return 0;
    p.drawPie(r, (int)(a1 * 16), (int)(span * 16));
    dc_done(p, g);
    return 1;
}
static vb_i64 u_DrawEdge(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    Gfx *g = gfx_of_dc((vb_i32)A(&a[0]));
    vb_i32 *r = (vb_i32 *)vb_dllarg_ptr(&a[1], 16);
    if (!g || !r) return 0;
    int edge = (int)A(&a[2]), fl = (int)A(&a[3]);
    QRect rr(QPoint(r[0], r[1]), QPoint(r[2] - 1, r[3] - 1));
    QPainter p(&g->img);
    QColor hi = qcolor(sys_color(20)), lt = qcolor(sys_color(22)), sh = qcolor(sys_color(16)), dk = qcolor(sys_color(21));
    auto side = [&](const QColor &tl, const QColor &br, QRect q) {
        p.setPen(tl);
        if (fl & 1) p.drawLine(q.topLeft(), q.bottomLeft());
        if (fl & 2) p.drawLine(q.topLeft(), q.topRight());
        p.setPen(br);
        if (fl & 4) p.drawLine(q.topRight(), q.bottomRight());
        if (fl & 8) p.drawLine(q.bottomLeft(), q.bottomRight());
    };
    if (edge & 0x3) side(edge & 1 ? lt : sh, edge & 1 ? dk : hi, rr);             /* the outer edge: raised / sunken */
    if (edge & 0xC) side(edge & 4 ? hi : dk, edge & 4 ? sh : lt, rr.adjusted(1, 1, -1, -1));
    if (fl & 0x800) p.fillRect(rr.adjusted(2, 2, -2, -2), qcolor(sys_color(15)));   /* BF_MIDDLE */
    p.end();
    gfx_update(g);
    return 1;
}
static vb_i64 g_FloodFill(vb_i32 argc, vb_dllarg *a, bool ext)
{
    if (argc < 4) return 0;
    vb_i32 dc = (vb_i32)A(&a[0]);
    Gfx *g = gfx_of_dc(dc);
    GdiObj *br = gdi_of(dc_state(dc).brush);
    if (!g) return 0;
    int x0 = (int)A(&a[1]), y0 = (int)A(&a[2]);
    QImage &img = g->img;
    if (x0 < 0 || y0 < 0 || x0 >= img.width() || y0 >= img.height()) return 0;
    QRgb key = qcolor((long)A(&a[3]) & 0xFFFFFF).rgb() & 0xFFFFFF, fill = (br ? br->color : QColor(Qt::white)).rgb();
    bool surface = ext && argc > 4 && A(&a[4]) == 1;   /* FLOODFILLSURFACE: the area of that color; else: up to that border */
    auto inside = [&](int x, int y) { QRgb c = img.pixel(x, y) & 0xFFFFFF; return surface ? c == key : c != key; };
    if (!inside(x0, y0) || (img.pixel(x0, y0) & 0xFFFFFF) == (fill & 0xFFFFFF)) return 1;
    std::vector<QPoint> st{ QPoint(x0, y0) };
    while (!st.empty()) {
        QPoint q = st.back(); st.pop_back();
        int x = q.x(), y = q.y();
        if (!inside(x, y) || (img.pixel(x, y) & 0xFFFFFF) == (fill & 0xFFFFFF)) continue;
        int l = x, r = x;
        while (l > 0 && inside(l - 1, y) && (img.pixel(l - 1, y) & 0xFFFFFF) != (fill & 0xFFFFFF)) l--;
        while (r < img.width() - 1 && inside(r + 1, y) && (img.pixel(r + 1, y) & 0xFFFFFF) != (fill & 0xFFFFFF)) r++;
        for (int i = l; i <= r; i++) {
            img.setPixel(i, y, fill);
            if (y > 0) st.push_back(QPoint(i, y - 1));
            if (y < img.height() - 1) st.push_back(QPoint(i, y + 1));
        }
    }
    gfx_update(g);
    return 1;
}
static vb_i64 g_FloodFill4(vb_i32 argc, vb_dllarg *a) { return g_FloodFill(argc, a, false); }
static vb_i64 g_ExtFloodFill(vb_i32 argc, vb_dllarg *a) { return g_FloodFill(argc, a, true); }
static vb_i64 gui_ret0(vb_i32, vb_dllarg *) { return 0; }
static vb_i64 gui_ret1(vb_i32, vb_dllarg *) { return 1; }
static vb_i64 gui_handle(vb_i32, vb_dllarg *) { static vb_i32 next = 0x00B00010; next += 4; return next; }
static vb_i64 u_RegisterWindowMessageA(vb_i32, vb_dllarg *) { static vb_i32 next = 0xC100; return next++; }
static vb_i64 u_GetWindowThreadProcessId(vb_i32 argc, vb_dllarg *a)
{
    if (argc > 1) { vb_i32 *p = (vb_i32 *)vb_dllarg_ptr(&a[1], 4); if (p && a[1].kind == VBD_PTR) *p = (vb_i32)QCoreApplication::applicationPid(); }
    return 1;
}
static vb_i64 u_MapVirtualKeyA(vb_i32 argc, vb_dllarg *a) { return argc ? A(&a[0]) & 0xFF : 0; }
static vb_i64 u_PeekMessageA(vb_i32, vb_dllarg *) { if (app) QCoreApplication::processEvents(); return 0; }
static vb_i64 u_GetMonitorInfoA(vb_i32 argc, vb_dllarg *a)
{
    vb_i32 *mi = argc > 1 ? (vb_i32 *)vb_dllarg_ptr(&a[1], 40) : nullptr;
    QScreen *s = QGuiApplication::primaryScreen();
    if (!mi || !s) return 0;
    QRect g = s->geometry(), av = s->availableGeometry();
    mi[1] = g.left(); mi[2] = g.top(); mi[3] = g.right() + 1; mi[4] = g.bottom() + 1;
    mi[5] = av.left(); mi[6] = av.top(); mi[7] = av.right() + 1; mi[8] = av.bottom() + 1;
    mi[9] = 1;                                  /* MONITORINFOF_PRIMARY */
    return 1;
}
static vb_i64 u_MonitorFromWindow(vb_i32, vb_dllarg *) { return 0x00C00001; }
static vb_i64 u_EnumDisplaySettingsA(vb_i32 argc, vb_dllarg *a)
{
    unsigned char *dm = argc > 2 ? (unsigned char *)vb_dllarg_ptr(&a[2], 124) : nullptr;
    QScreen *s = QGuiApplication::primaryScreen();
    vb_i32 mode = argc > 1 ? (vb_i32)A(&a[1]) : 0;
    if (!dm || !s || (mode >= 1 && mode != -1 && mode != -2)) return 0;   /* the current mode only */
    vb_i32 bpp = s->depth(), w = s->geometry().width(), h = s->geometry().height(), hz = (vb_i32)s->refreshRate();
    memcpy(dm + 104, &bpp, 4); memcpy(dm + 108, &w, 4); memcpy(dm + 112, &h, 4); memcpy(dm + 120, &hz, 4);
    return 1;
}
static vb_i64 u_GetClassLongA(vb_i32 argc, vb_dllarg *a)
{
    int idx = argc > 1 ? (int)A(&a[1]) : 0;
    return idx == -26 ? 0x0008 : 0;            /* GCL_STYLE: CS_DBLCLKS */
}
static vb_i64 u_MessageBoxIndirectA(vb_i32 argc, vb_dllarg *a)
{
    unsigned char *p = argc ? (unsigned char *)vb_dllarg_ptr(&a[0], 40) : nullptr;
    if (!p) return 0;
    vb_i32 flags; memcpy(&flags, p + 20, 4);
    vb_i32 mk = vb_mark();
    vb_i32 r = vb_msgbox_hook ? vb_msgbox_hook(vs(ansi_q(ptr_str(p, 12))), flags & 0xFFFF, vs(ansi_q(ptr_str(p, 16)))) : 1;
    vb_release(mk);
    return r;
}

/* the size of text in the font of a device context */
static vb_i64 g_GetTextExtentPoint32A(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    QString t = astr(&a[1]);
    if (A(&a[2]) >= 0) t = t.left((int)A(&a[2]));
    vb_i32 *sz = (vb_i32 *)vb_dllarg_ptr(&a[3], 8);
    if (!sz) return 0;
    QFontMetrics fm(dc_font((vb_i32)A(&a[0])));
    sz[0] = fm.horizontalAdvance(t); sz[1] = fm.height();
    return 1;
}
static vb_i64 g_GetTextMetricsA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    unsigned char *tm = (unsigned char *)vb_dllarg_ptr(&a[1], 53);
    if (!tm) return 0;
    QFont f = dc_font((vb_i32)A(&a[0]));
    QFontMetrics fm(f);
    vb_i32 v[8] = { fm.height(), fm.ascent(), fm.descent(), fm.leading() > 0 ? 0 : 0, fm.leading(),
                    fm.averageCharWidth(), fm.maxWidth(), f.bold() ? 700 : 400 };
    memset(tm, 0, 53);
    memcpy(tm, v, sizeof v);
    tm[44] = ' '; tm[45] = 255; tm[46] = '?'; tm[47] = ' ';   /* first, last, default, break character */
    tm[48] = f.italic(); tm[49] = f.underline(); tm[50] = f.strikeOut(); tm[51] = 0x26;   /* pitch and family */
    return 1;
}

static vb_i64 g_FrameRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 5) return 0;
    Gfx *g = gfx_of_dc((vb_i32)A(&a[0]));
    GdiObj *r = gdi_of((vb_i32)A(&a[1])), *b = gdi_of((vb_i32)A(&a[2]));
    if (!g || !r) return 0;
    int w = std::max(1, (int)A(&a[3])), h = std::max(1, (int)A(&a[4]));
    QRegion inner = r->rgn.intersected(r->rgn.translated(w, 0)).intersected(r->rgn.translated(-w, 0))
                           .intersected(r->rgn.translated(0, h)).intersected(r->rgn.translated(0, -h));
    QRegion edge = r->rgn.subtracted(inner);
    QPainter p(&g->img);
    p.setClipRegion(edge);
    p.fillRect(edge.boundingRect(), b ? b->color : QColor(Qt::black));
    p.end();
    gfx_update(g);
    return 1;
}
static vb_i64 g_InvertRgn(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    Gfx *g = gfx_of_dc((vb_i32)A(&a[0]));
    GdiObj *r = gdi_of((vb_i32)A(&a[1]));
    if (!g || !r) return 0;
    QPainter p(&g->img);
    p.setClipRegion(r->rgn);
    p.setCompositionMode(QPainter::RasterOp_SourceXorDestination);
    p.fillRect(r->rgn.boundingRect(), Qt::white);
    p.end();
    gfx_update(g);
    return 1;
}
static vb_i64 g_GetRgnBox(vb_i32 argc, vb_dllarg *a)
{
    GdiObj *r = argc >= 2 ? gdi_of((vb_i32)A(&a[0])) : nullptr;
    vb_i32 *rc = argc >= 2 ? (vb_i32 *)vb_dllarg_ptr(&a[1], 16) : nullptr;
    if (!r || !rc) return 0;
    QRect b = r->rgn.boundingRect();
    rc[0] = b.left(); rc[1] = b.top(); rc[2] = b.right() + 1; rc[3] = b.bottom() + 1;
    return r->rgn.isEmpty() ? 1 : r->rgn.rectCount() == 1 ? 2 : 3;
}

struct ApiName { const char *n; vb_apiext f; };
/* ---------------------------------------------------------------- more of the Planet Source Code programs */
/* msimg32 GradientFill: TRIVERTEX { x, y As Long; Red, Green, Blue, Alpha As Integer (8.8) }; rectangles
 * (GRADIENT_FILL_RECT_H 0, _V 1) and triangles (2) */
static vb_i64 m_GradientFill(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 6) return 0;
    int nv = (int)A(&a[2]), nm = (int)A(&a[4]), mode = (int)A(&a[5]);
    if (nv <= 0 || nm <= 0 || nv > 100000 || nm > 100000) return 0;
    unsigned char *v = (unsigned char *)vb_dllarg_ptr(&a[1], nv * 16);
    vb_i32 *m = (vb_i32 *)vb_dllarg_ptr(&a[3], nm * (mode == 2 ? 12 : 8));
    QPainter p; Gfx *g;
    if (!v || !m || !dc_painter((vb_i32)A(&a[0]), p, &g)) return 0;
    auto vx = [&](int i, int *x, int *y, QColor *c) {
        vb_i32 xy[2]; vb_u16 rgba[4];
        memcpy(xy, v + i * 16, 8); memcpy(rgba, v + i * 16 + 8, 8);
        *x = xy[0]; *y = xy[1];
        *c = QColor(rgba[0] >> 8, rgba[1] >> 8, rgba[2] >> 8);
    };
    p.setPen(Qt::NoPen);
    for (int k = 0; k < nm; k++) {
        if (mode == 2) {                        /* a triangle: each pixel weighs the colors of the corners */
            int x[3], y[3]; QColor c[3];
            bool ok = true;
            for (int j = 0; j < 3; j++) { int i = m[k * 3 + j]; if (i < 0 || i >= nv) { ok = false; break; } vx(i, &x[j], &y[j], &c[j]); }
            if (!ok) continue;
            double den = (double)(y[1] - y[2]) * (x[0] - x[2]) + (double)(x[2] - x[1]) * (y[0] - y[2]);
            if (den == 0) continue;
            int x0 = std::max(0, std::min({ x[0], x[1], x[2] })), x1 = std::min(g->img.width() - 1, std::max({ x[0], x[1], x[2] }));
            int y0 = std::max(0, std::min({ y[0], y[1], y[2] })), y1 = std::min(g->img.height() - 1, std::max({ y[0], y[1], y[2] }));
            p.end();
            for (int yy = y0; yy <= y1; yy++)
                for (int xx = x0; xx <= x1; xx++) {
                    double w0 = ((y[1] - y[2]) * (xx - x[2]) + (x[2] - x[1]) * (yy - y[2])) / den;
                    double w1 = ((y[2] - y[0]) * (xx - x[2]) + (x[0] - x[2]) * (yy - y[2])) / den;
                    double w2 = 1 - w0 - w1;
                    if (w0 < 0 || w1 < 0 || w2 < 0) continue;
                    g->img.setPixel(xx, yy, qRgb((int)(w0 * c[0].red() + w1 * c[1].red() + w2 * c[2].red()),
                                                 (int)(w0 * c[0].green() + w1 * c[1].green() + w2 * c[2].green()),
                                                 (int)(w0 * c[0].blue() + w1 * c[1].blue() + w2 * c[2].blue())));
                }
            p.begin(&g->img);
            continue;
        }
        int i0 = m[k * 2], i1 = m[k * 2 + 1];
        if (i0 < 0 || i0 >= nv || i1 < 0 || i1 >= nv) continue;
        int xa, ya, xb, yb; QColor ca, cb;
        vx(i0, &xa, &ya, &ca); vx(i1, &xb, &yb, &cb);
        QLinearGradient gr(mode == 0 ? QPointF(xa, 0) : QPointF(0, ya), mode == 0 ? QPointF(xb, 0) : QPointF(0, yb));
        gr.setColorAt(0, ca); gr.setColorAt(1, cb);
        p.fillRect(QRect(QPoint(std::min(xa, xb), std::min(ya, yb)), QPoint(std::max(xa, xb) - 1, std::max(ya, yb) - 1)), gr);
    }
    dc_done(p, g);
    return 1;
}
/* GetBitmapBits / SetBitmapBits: the pixels of a bitmap as 32-bit B G R x */
static vb_i64 g_BitmapBits(vb_i32 argc, vb_dllarg *a, bool put)
{
    if (argc < 3) return 0;
    GdiObj *o = gdi_of((vb_i32)A(&a[0]));
    Gfx *b = o && o->type == 3 ? o->bmp : memdc_gfx((vb_i32)A(&a[0]));
    vb_i64 n = A(&a[1]);
    if (!b || n <= 0) return 0;
    QImage &img = b->img;
    if (img.format() != QImage::Format_RGB32) img = img.convertToFormat(QImage::Format_RGB32);
    vb_i64 total = (vb_i64)img.width() * img.height() * 4;
    if (n > total) n = total;
    unsigned char *buf = (unsigned char *)vb_dllarg_ptr(&a[2], (vb_i32)n);
    if (!buf) return 0;
    vb_i64 row = (vb_i64)img.width() * 4;
    for (vb_i64 off = 0; off < n; off += row) {
        vb_i64 k = std::min(row, n - off);
        unsigned char *line = img.scanLine((int)(off / row));
        if (put) memcpy(line, buf + off, (size_t)k); else memcpy(buf + off, line, (size_t)k);
    }
    if (put) gfx_update(b);
    return n;
}
static vb_i64 g_GetBitmapBits(vb_i32 argc, vb_dllarg *a) { return g_BitmapBits(argc, a, false); }
static vb_i64 g_SetBitmapBits(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    vb_dllarg b[3] = { a[0], a[1], a[2] };
    return g_BitmapBits(3, b, true);
}
/* PlgBlt: a rectangle of the source onto a parallelogram (three corners: upper left, upper right, lower left) */
static vb_i64 g_PlgBlt(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 7) return 0;
    vb_i32 *pt = (vb_i32 *)vb_dllarg_ptr(&a[1], 24);
    int sw = (int)A(&a[5]), sh = (int)A(&a[6]), ox, oy;
    if (!pt || sw <= 0 || sh <= 0) return 0;
    QImage src = source_image((vb_i32)A(&a[2]), (int)A(&a[3]), (int)A(&a[4]), sw, sh, &ox, &oy);
    if (src.isNull()) return 0;
    QPainter p; Gfx *g;
    if (!dc_painter((vb_i32)A(&a[0]), p, &g)) return 0;
    QTransform t(double(pt[2] - pt[0]) / sw, double(pt[3] - pt[1]) / sw,
                 double(pt[4] - pt[0]) / sh, double(pt[5] - pt[1]) / sh, pt[0], pt[1]);
    p.setTransform(t);
    p.setRenderHint(QPainter::SmoothPixmapTransform);
    p.drawImage(QPoint(0, 0), src, QRect(ox, oy, sw, sh));
    dc_done(p, g);
    return 1;
}
/* ExtCreateRegion: RGNDATA = RGNDATAHEADER (32 bytes: dwSize, iType, nCount, nRgnSize, rcBound) + nCount RECTs */
static vb_i64 g_ExtCreateRegion(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 3) return 0;
    vb_i64 size = A(&a[1]);
    vb_i32 *h = (vb_i32 *)vb_dllarg_ptr(&a[2], 32);
    if (!h || h[2] < 0) return 0;
    vb_i32 *r = (vb_i32 *)vb_dllarg_ptr(&a[2], 32 + h[2] * 16);
    if (!r || (size && size < 32 + (vb_i64)h[2] * 16)) return 0;
    r += 8;
    QRegion rgn;
    for (int i = 0; i < h[2]; i++, r += 4) rgn += QRect(QPoint(r[0], r[1]), QPoint(r[2] - 1, r[3] - 1));
    return rgn_new(rgn);
}
/* GetWindowPlacement / SetWindowPlacement: WINDOWPLACEMENT (44 bytes) */
static vb_i64 u_GetWindowPlacement(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    QWidget *w = widget_of_hwnd((vb_i32)A(&a[0]));
    vb_i32 *wp = (vb_i32 *)vb_dllarg_ptr(&a[1], 44);
    if (!w || !wp) return 0;
    QWidget *t = w->window();
    QRect g = t->normalGeometry().isValid() ? t->normalGeometry() : t->frameGeometry();
    wp[0] = 44; wp[1] = 0;
    wp[2] = !t->isVisible() ? 0 : t->isMinimized() ? 2 : t->isMaximized() ? 3 : 1;   /* SW_HIDE, SHOWMINIMIZED ... */
    wp[3] = wp[4] = -1; wp[5] = wp[6] = -1;
    wp[7] = g.left(); wp[8] = g.top(); wp[9] = g.right() + 1; wp[10] = g.bottom() + 1;
    return 1;
}
static vb_i64 u_SetWindowPlacement(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    QWidget *w = widget_of_hwnd((vb_i32)A(&a[0]));
    vb_i32 *wp = (vb_i32 *)vb_dllarg_ptr(&a[1], 44);
    if (!w || !wp) return 0;
    QWidget *t = w->window();
    t->setGeometry(QRect(QPoint(wp[7], wp[8]), QPoint(wp[9] - 1, wp[10] - 1)));
    if (wp[2] == 2 || wp[2] == 6 || wp[2] == 7) t->showMinimized();
    else if (wp[2] == 3) t->showMaximized();
    else if (wp[2] == 0) t->hide();
    else t->showNormal();
    return 1;
}
/* SetPropA / GetPropA / RemovePropA: values kept by window and name (or atom) */
static std::map<std::pair<vb_i32, QString>, vb_i64> wprops;
static QString prop_key(vb_dllarg *a) { return a->kind == VBD_STR ? astr(a).toLower() : QString("#%1").arg(A(a)); }
static vb_i64 u_SetPropA(vb_i32 argc, vb_dllarg *a) { if (argc < 3) return 0; wprops[{ (vb_i32)A(&a[0]), prop_key(&a[1]) }] = A(&a[2]); return 1; }
static vb_i64 u_GetPropA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    auto it = wprops.find({ (vb_i32)A(&a[0]), prop_key(&a[1]) });
    return it == wprops.end() ? 0 : it->second;
}
static vb_i64 u_RemovePropA(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    auto it = wprops.find({ (vb_i32)A(&a[0]), prop_key(&a[1]) });
    if (it == wprops.end()) return 0;
    vb_i64 v = it->second; wprops.erase(it); return v;
}
/* DrawFrameControl: the face and the 3-D edges of a button (the frame types are drawn alike) */
static vb_i64 u_DrawFrameControl(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return 0;
    vb_i32 *r = (vb_i32 *)vb_dllarg_ptr(&a[1], 16);
    QPainter p; Gfx *g;
    if (!r || !dc_painter((vb_i32)A(&a[0]), p, &g)) return 0;
    QRect rr(QPoint(r[0], r[1]), QPoint(r[2] - 1, r[3] - 1));
    bool pushed = (A(&a[3]) & 0x0200) != 0;      /* DFCS_PUSHED */
    p.fillRect(rr, qcolor(sys_color(15)));
    p.setPen(qcolor(sys_color(pushed ? 16 : 20)));
    p.drawLine(rr.topLeft(), rr.topRight()); p.drawLine(rr.topLeft(), rr.bottomLeft());
    p.setPen(qcolor(sys_color(pushed ? 20 : 16)));
    p.drawLine(rr.bottomLeft(), rr.bottomRight()); p.drawLine(rr.topRight(), rr.bottomRight());
    dc_done(p, g);
    return 1;
}
static vb_i64 g_GetTextColor(vb_i32 argc, vb_dllarg *a) { if (argc < 1) return 0; return vbcolor(dc_state((vb_i32)A(&a[0])).text.rgb()); }
static vb_i64 g_GetBkColor(vb_i32 argc, vb_dllarg *a) { if (argc < 1) return 0; return vbcolor(dc_state((vb_i32)A(&a[0])).back.rgb()); }
static vb_i64 g_GetColorAdjustment(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 2) return 0;
    vb_i16 *ca = (vb_i16 *)vb_dllarg_ptr(&a[1], 24);
    if (!ca) return 0;
    memset(ca, 0, 24);
    ca[0] = 24; ca[3] = ca[4] = ca[5] = 10000; ca[7] = 10000;     /* caSize, gammas, reference white */
    return 1;
}
static vb_i64 one_true(vb_i32, vb_dllarg *) { return 1; }
static vb_i64 zero_false(vb_i32, vb_dllarg *) { return 0; }
/* winmm timeSetEvent / timeKillEvent: the callback (AddressOf) on the program's thread, by a QTimer */
static std::map<vb_i32, QTimer *> mm_timers;
static vb_i64 w_timeSetEvent(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 5) return 0;
    static vb_i32 next_id = 1;
    int delay = (int)A(&a[0]);
    vb_i64 fn = A(&a[2]), user = A(&a[3]), flags = A(&a[4]);
    if (!vb_callback_fn(fn)) return 0;
    vb_i32 id = next_id++;
    auto *t = new QTimer;
    t->setSingleShot(!(flags & 1));             /* TIME_PERIODIC */
    t->setTimerType(Qt::PreciseTimer);
    QObject::connect(t, &QTimer::timeout, [id, fn, user, t]() {
        vb_i32 args[5] = { id, 0, (vb_i32)user, 0, 0 };
        vb_callback(fn, 5, args);
        if (t->isSingleShot()) { mm_timers.erase(id); t->deleteLater(); }
    });
    mm_timers[id] = t;
    t->start(std::max(1, delay));
    return id;
}
static vb_i64 w_timeKillEvent(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 1) return 97;                     /* MMSYSERR_INVALPARAM */
    auto it = mm_timers.find((vb_i32)A(&a[0]));
    if (it == mm_timers.end()) return 97;
    it->second->stop(); it->second->deleteLater(); mm_timers.erase(it);
    return 0;
}

static const ApiName gui_user32[] = {
    { "GetWindowPlacement", u_GetWindowPlacement }, { "SetWindowPlacement", u_SetWindowPlacement },
    { "SetPropA", u_SetPropA }, { "GetPropA", u_GetPropA }, { "RemovePropA", u_RemovePropA },
    { "DrawFrameControl", u_DrawFrameControl }, { "DrawAnimatedRects", one_true }, { "GetInputState", zero_false },
    { "GetWindowRgn", zero_false },             /* ERROR: the window has no region of its own */
    { "CreateCaret", one_true }, { "ShowCaret", one_true }, { "HideCaret", one_true }, { "DestroyCaret", one_true },
    { "SetCaretPos", one_true }, { "ChangeDisplaySettingsA", zero_false }, { "ChangeDisplaySettings", zero_false },
    { "GetAsyncKeyState", u_GetAsyncKeyState }, { "GetKeyState", u_GetKeyState }, { "GetKeyboardState", u_GetKeyboardState },
    { "keybd_event", u_keybd_event },
    { "FindWindowA", u_FindWindowA }, { "FindWindowExA", u_FindWindowExA },
    { "GetWindowTextA", u_GetWindowTextA }, { "GetWindowTextLengthA", u_GetWindowTextLengthA }, { "SetWindowTextA", u_SetWindowTextA },
    { "ShowWindow", u_ShowWindow }, { "IsWindow", u_IsWindow }, { "IsWindowVisible", u_IsWindowVisible },
    { "IsWindowEnabled", u_IsWindowEnabled }, { "IsIconic", u_IsIconic }, { "IsZoomed", u_IsZoomed },
    { "EnableWindow", u_EnableWindow }, { "MoveWindow", u_MoveWindow }, { "BringWindowToTop", u_BringWindowToTop },
    { "SetForegroundWindow", u_BringWindowToTop }, { "SetActiveWindow", u_BringWindowToTop },
    { "GetParent", u_GetParent }, { "SetParent", u_SetParent }, { "DestroyWindow", u_DestroyWindow },
    { "CloseWindow", u_CloseWindow }, { "FlashWindow", u_FlashWindow }, { "FlashWindowEx", u_FlashWindow },
    { "GetWindow", u_GetWindow }, { "GetClassNameA", u_GetClassNameA },
    { "WindowFromPoint", u_WindowFromPoint }, { "ChildWindowFromPoint", u_WindowFromPoint },
    { "GetFocus", u_GetFocus }, { "SetFocus", u_SetFocus }, { "SetCapture", u_SetCapture }, { "GetCapture", u_GetCapture },
    { "ReleaseCapture", u_ReleaseCapture }, { "InvalidateRect", u_InvalidateRect }, { "UpdateWindow", u_UpdateWindow },
    { "RedrawWindow", u_InvalidateRect }, { "LockWindowUpdate", u_LockWindowUpdate },
    { "EnumWindows", u_EnumWindows }, { "EnumChildWindows", u_EnumChildWindows },
    { "GetWindowLongA", u_GetWindowLongA }, { "SetWindowLongA", u_SetWindowLongA },
    { "SetLayeredWindowAttributes", u_SetLayeredWindowAttributes },
    { "GetSystemMenu", u_GetSystemMenu }, { "GetMenuItemID", u_GetMenuItemID },
    { "ShowCursor", u_ShowCursor }, { "SystemParametersInfoA", u_SystemParametersInfoA }, { "GetSysColor", u_GetSysColor },
    { "SendMessageA", u_SendMessageA }, { "SendMessageLongA", u_SendMessageA }, { "SendMessageStringA", u_SendMessageA },
    { "SendMessageLong", u_SendMessageA }, { "SendMessageString", u_SendMessageA },
    { "PostMessageA", u_PostMessageA }, { "SendMessageTimeoutA", u_SendMessageA },
    { "FillRect", u_FillRect }, { "FrameRect", u_FrameRect }, { "DrawTextA", u_DrawTextA }, { "DrawTextExA", u_DrawTextA },
    { "SetWindowRgn", u_SetWindowRgn }, { "LoadImageA", u_LoadImageA }, { "LoadBitmapA", u_LoadImageA },
    { "DrawEdge", u_DrawEdge }, { "RegisterWindowMessageA", u_RegisterWindowMessageA },
    { "GetWindowThreadProcessId", u_GetWindowThreadProcessId }, { "MapVirtualKeyA", u_MapVirtualKeyA },
    { "PeekMessageA", u_PeekMessageA }, { "GetMonitorInfoA", u_GetMonitorInfoA }, { "MonitorFromWindow", u_MonitorFromWindow },
    { "MonitorFromPoint", u_MonitorFromWindow }, { "EnumDisplaySettingsA", u_EnumDisplaySettingsA },
    { "GetClassLongA", u_GetClassLongA }, { "SetClassLongA", gui_ret0 }, { "SetWindowWord", gui_ret0 },
    { "GetWindowWord", gui_ret0 }, { "MessageBoxIndirectA", u_MessageBoxIndirectA },
    { "CreatePopupMenu", gui_handle }, { "CreateMenu", gui_handle }, { "DestroyMenu", gui_ret1 }, { "IsMenu", gui_ret1 },
    { "SetMenuItemBitmaps", gui_ret1 }, { "InsertMenuItemA", gui_ret1 }, { "GetMenuItemInfoA", gui_ret0 },
    { "SetMenuItemInfoA", gui_ret1 }, { "GetMenuItemRect", gui_ret0 }, { "RemoveMenu", gui_ret1 }, { "DeleteMenu", gui_ret1 },
    { "AppendMenuA", gui_ret1 }, { "InsertMenuA", gui_ret1 }, { "ModifyMenuA", gui_ret1 }, { "DrawMenuBar", gui_ret1 },
    { "CheckMenuItem", gui_ret0 }, { "EnableMenuItem", gui_ret0 }, { "SetMenuDefaultItem", gui_ret1 },
    { "LoadCursorFromFileA", gui_handle }, { "LoadCursorA", gui_handle }, { "LoadIconA", gui_handle },
    { "GetCursor", gui_handle }, { "SetCursor", gui_ret0 }, { "SetSystemCursor", gui_ret1 }, { "DestroyCursor", gui_ret1 },
    { "DestroyIcon", gui_ret1 }, { "DrawIconEx", gui_ret1 }, { "DrawIcon", gui_ret1 }, { "SetKeyboardState", gui_ret1 },
    { "GetQueueStatus", gui_ret0 }, { "PaintDesktop", gui_ret1 }, { "BlockInput", gui_ret1 }, { "SetDoubleClickTime", gui_ret1 },
    { "GetClipCursor", gui_ret1 }, { "DrawFocusRect", gui_ret1 }, { "InvertRect", gui_ret1 }, { "DrawStateA", gui_ret1 },
    { "CreateWindowExA", gui_ret0 }, { "AnimateWindow", u_ShowWindow }, { "UpdateLayeredWindow", gui_ret1 },
};
static const ApiName gui_gdi32[] = {
    { "GetBitmapBits", g_GetBitmapBits }, { "SetBitmapBits", g_SetBitmapBits }, { "PlgBlt", g_PlgBlt },
    { "ExtCreateRegion", g_ExtCreateRegion }, { "GetTextColor", g_GetTextColor }, { "GetBkColor", g_GetBkColor },
    { "GetColorAdjustment", g_GetColorAdjustment }, { "SetColorAdjustment", one_true }, { "SetTextJustification", one_true },
    { "CreateCompatibleDC", g_CreateCompatibleDC }, { "CreateDCA", g_CreateDCA }, { "DeleteDC", g_DeleteDC },
    { "CreateCompatibleBitmap", g_CreateCompatibleBitmap }, { "CreateBitmap", g_CreateBitmap }, { "CreatePen", g_CreatePen },
    { "CreateSolidBrush", g_CreateSolidBrush }, { "CreateHatchBrush", g_CreateHatchBrush }, { "GetStockObject", g_GetStockObject },
    { "CreatePatternBrush", g_CreatePatternBrush }, { "Escape", zero_false },
    { "SelectObject", g_SelectObject }, { "DeleteObject", g_DeleteObject }, { "GetObjectA", g_GetObjectA },
    { "SetTextColor", g_SetTextColor }, { "SetBkColor", g_SetBkColor }, { "SetBkMode", g_SetBkMode },
    { "SetStretchBltMode", g_SetStretchBltMode }, { "MoveToEx", g_MoveToEx }, { "LineTo", g_LineTo },
    { "Rectangle", g_Rectangle }, { "Ellipse", g_Ellipse }, { "RoundRect", g_RoundRect },
    { "Polygon", g_Polygon }, { "Polyline", g_Polyline }, { "PolyBezier", g_PolyBezier }, { "TextOutA", g_TextOutA },
    { "CreateFontA", g_CreateFontA },
    { "CreateRectRgn", g_CreateRectRgn }, { "CreateEllipticRgn", g_CreateEllipticRgn },
    { "CreateRoundRectRgn", g_CreateRoundRectRgn }, { "CreatePolygonRgn", g_CreatePolygonRgn },
    { "CreatePolyPolygonRgn", g_CreatePolyPolygonRgn }, { "CombineRgn", g_CombineRgn }, { "OffsetRgn", g_OffsetRgn },
    { "PtInRegion", g_PtInRegion }, { "FillRgn", g_FillRgn },
    { "FrameRgn", g_FrameRgn }, { "InvertRgn", g_InvertRgn }, { "PaintRgn", g_FillRgn }, { "GetRgnBox", g_GetRgnBox },
    { "GetTextExtentPoint32A", g_GetTextExtentPoint32A }, { "GetTextExtentPointA", g_GetTextExtentPoint32A },
    { "GetTextMetricsA", g_GetTextMetricsA },
    { "GetDIBits", g_GetDIBits }, { "SetDIBits", g_SetDIBits }, { "StretchDIBits", g_StretchDIBits },
    { "SetDIBitsToDevice", g_SetDIBitsToDevice }, { "CreateBrushIndirect", g_CreateBrushIndirect },
    { "CreateFontIndirectA", g_CreateFontIndirectA }, { "CreateFontIndirect", g_CreateFontIndirectA }, { "Pie", g_Pie },
    { "FloodFill", g_FloodFill4 }, { "ExtFloodFill", g_ExtFloodFill },
    { "SetGraphicsMode", gui_ret1 }, { "SetMapMode", gui_ret1 }, { "GetPolyFillMode", gui_ret1 }, { "SetPolyFillMode", gui_ret1 },
    { "GdiTransparentBlt", m_TransparentBlt }, { "SetROP2", gui_ret1 }, { "GetDeviceGammaRamp", gui_ret0 },
    { "SetDeviceGammaRamp", gui_ret0 }, { "AddFontResourceA", gui_ret1 }, { "RemoveFontResourceA", gui_ret1 },
    { "BeginPath", gui_ret1 }, { "EndPath", gui_ret1 }, { "SelectClipRgn", gui_ret1 }, { "GdiFlush", gui_ret1 },
};
/* OleCreatePictureIndirect(PICTDESC, riid, fOwn, IPic): a StdPicture of a bitmap (or an icon) handle */
static vb_i64 o_OleCreatePictureIndirect(vb_i32 argc, vb_dllarg *a)
{
    if (argc < 4) return (vb_i32)0x80070057;     /* E_INVALIDARG */
    unsigned char *pd = (unsigned char *)vb_dllarg_ptr(&a[0], 12);
    if (!pd) return (vb_i32)0x80070057;
    vb_i32 type, h;
    memcpy(&type, pd + 4, 4); memcpy(&h, pd + 8, 4);
    QImage img;
    Gfx *g = bitmap_gfx(h);
    if (g) img = g->img.copy();
    else { auto it = pic_handles.find(h); if (it != pic_handles.end() && it->second) img = it->second->img; }
    vbobj pic = new_pic(img, type == 3 ? 3 : 1);
    vb_oaddref(pic);                            /* the object variable's reference */
    if (a[3].kind == VBD_PTR && a[3].p) *(vb_i64 *)a[3].p = vb_varptr(pic, (vb_i32)sizeof(vbobjh));
    return 0;                                   /* S_OK */
}
static const ApiName gui_other[] = {
    { "OleCreatePictureIndirect", o_OleCreatePictureIndirect },
    { "Shell_NotifyIconA", s_Shell_NotifyIconA }, { "Shell_NotifyIcon", s_Shell_NotifyIconA },
#ifndef _WIN32
    { "SHBrowseForFolderA", s_SHBrowseForFolderA }, { "SHBrowseForFolder", s_SHBrowseForFolderA },
#endif
    { "GetOpenFileNameA", c_GetOpenFileNameA }, { "GetSaveFileNameA", c_GetSaveFileNameA },
    { "ChooseColorA", c_ChooseColorA }, { "CommDlgExtendedError", c_CommDlgExtendedError },
    { "TransparentBlt", m_TransparentBlt }, { "AlphaBlend", m_AlphaBlend }, { "OleTranslateColor", o_OleTranslateColor },
    { "GradientFill", m_GradientFill }, { "timeSetEvent", w_timeSetEvent }, { "timeKillEvent", w_timeKillEvent },
};
static vb_apiext gui_api(const char *lib, const char *name)
{
#ifdef _WIN32
    /* on Windows the real procedures where no window or device context of the program is involved */
    static const char *const real[] = { "GetAsyncKeyState", "GetKeyState", "GetKeyboardState", "SetKeyboardState",
        "keybd_event", "GetSysColor", "SystemParametersInfoA", "ShowCursor", "OleTranslateColor", "GetMonitorInfoA",
        "MonitorFromPoint", "EnumDisplaySettingsA", "RegisterWindowMessageA", "MapVirtualKeyA", "AddFontResourceA",
        "RemoveFontResourceA", "SetSystemCursor", "LoadCursorFromFileA", "LoadCursorA", "LoadIconA", "GetCursor",
        "SetCursor", "DestroyCursor", "DestroyIcon", "BlockInput", "SetDoubleClickTime", "GetClipCursor",
        "GetQueueStatus", "GetDeviceGammaRamp", "SetDeviceGammaRamp", "PaintDesktop", "timeSetEvent", "timeKillEvent",
        "GetInputState", "ChangeDisplaySettingsA", "ChangeDisplaySettings", "CreateCaret", "ShowCaret", "HideCaret",
        "DestroyCaret", "SetCaretPos", 0 };
    for (int i = 0; real[i]; i++) if (!strcmp(real[i], name)) return nullptr;
#endif
    const ApiName *t = nullptr; size_t n = 0;
    if (!strcmp(lib, "user32")) { t = gui_user32; n = sizeof gui_user32 / sizeof *gui_user32; }
    else if (!strcmp(lib, "gdi32")) { t = gui_gdi32; n = sizeof gui_gdi32 / sizeof *gui_gdi32; }
    else { t = gui_other; n = sizeof gui_other / sizeof *gui_other; }
    for (size_t i = 0; i < n; i++)
        if (!strcmp(t[i].n, name)) { ensure_app(); return t[i].f; }   /* (ShowCursor in Sub Main: Qt first) */
    return nullptr;
}
