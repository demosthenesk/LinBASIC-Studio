"""The Toolbar Wizard (phase 9.2a-4): Add-Ins > Toolbar Wizard... when the "VB 6 Application Wizard" add-in is
loaded (VB6 oracle DEV-9/aw20 - aw22).

Three pages, as VB6: Introduction (a profile to load), Customize Toolbar (the toolbar list empty, all the buttons
on the left), Finished! (a profile to save). Finish: on the form of the active designer, Toolbar1 (Align Top) and,
with buttons, imlToolbarIcons (ImageList 16 x 16, MaskColor &HC0C0C0) and Toolbar1_ButtonClick (Select Case
Button.Key, a TODO for each button); without buttons only Toolbar1 (aw20). The wizard does not open by itself
when a Toolbar is placed on a form (aw22). The wording is our own.
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import os

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QStackedWidget, QFrame,
                             QMessageBox)

from . import appwizgen as G
from .appwizard import ToolbarEditor, wizard_picture, NONE, AppWizard

TITLE = "Toolbar Wizard"
PAGES = ("Introduction", "Customize Toolbar", "Finished!")


class ToolbarWizard(QDialog):
    def __init__(self, main, designer):
        super().__init__(main)
        self.main = main
        self.designer = designer
        self.buttons = []
        self.images = {}
        self.made = None                        # the name of the Toolbar, after Finish
        outer = QVBoxLayout(self)
        body = QHBoxLayout()
        self.picture = QLabel()
        self.picture.setFixedSize(150, 120)
        self.picture.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(self.picture, 0, Qt.AlignmentFlag.AlignTop)
        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        outer.addLayout(body, 1)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        outer.addWidget(line)
        row = QHBoxLayout()
        b = QPushButton("Help")
        b.clicked.connect(lambda: main.help_keyword("Toolbar Wizard"))
        row.addWidget(b)
        row.addStretch(1)
        self.b_cancel = QPushButton("Cancel")
        self.b_back = QPushButton("< &Back")
        self.b_next = QPushButton("&Next >")
        self.b_finish = QPushButton("&Finish")
        for b in (self.b_cancel, self.b_back, self.b_next, self.b_finish):
            row.addWidget(b)
        self.b_cancel.clicked.connect(self.reject)
        self.b_back.clicked.connect(lambda: self.go(self.stack.currentIndex() - 1))
        self.b_next.clicked.connect(lambda: self.go(self.stack.currentIndex() + 1))
        self.b_finish.clicked.connect(self.finish)
        outer.addLayout(row)
        # Introduction
        lay = self._page("The Toolbar Wizard makes a new toolbar on the form, with the buttons you choose and "
                         "their pictures.\n\nYou can go Back at any time to change your choices.\n\nClick Next to "
                         "begin.")
        lay.addStretch(1)
        self.load_combo = AppWizard._profile_row(self, lay, "From what profile do you want to load your settings?",
                                                 False)
        self.load_combo.activated.connect(lambda _i: self.load_profile(self.load_combo.currentText()))
        # Customize Toolbar
        lay = self._page("Make the toolbar: move the buttons you want to the list on the right, put them in order "
                         "with the arrows, and add pictures of your own with the picture button. A double click "
                         "moves a button too.")
        self.tbed = ToolbarEditor(self, self.images, reset=self.reset)
        lay.addWidget(self.tbed, 1)
        self.tbed.set_buttons(self.buttons)
        # Finished!
        lay = self._page("The Toolbar Wizard has all the information it needs.\n\nClick Finish to build your "
                         "toolbar.")
        lay.addStretch(1)
        self.save_combo = AppWizard._profile_row(self, lay, "To what profile do you want to save your settings?",
                                                 True)
        lay.addStretch(1)
        self.resize(720, 470)
        self.go(0)

    _page = AppWizard._page

    def pick_profile(self, combo, save):
        AppWizard.pick_profile(self, combo, save)

    def go(self, i):
        i = max(0, min(i, len(PAGES) - 1))
        self.stack.setCurrentIndex(i)
        self.setWindowTitle(f"{TITLE} - {PAGES[i]}")
        self.picture.setPixmap(wizard_picture(("intro", "toolbar", "finished")[i]))
        self.b_back.setEnabled(i > 0)
        self.b_next.setEnabled(i < len(PAGES) - 1)
        self.b_finish.setEnabled(i > 0)
        (self.b_finish if i == len(PAGES) - 1 else self.b_next).setDefault(True)

    def reset(self):
        del self.buttons[:]
        self.tbed.fill()

    def load_profile(self, path):
        if not path or path == NONE:
            return
        try:
            with open(path, encoding="cp1252") as f:
                o = G.AppOptions.from_rwp(f.read())
        except OSError as e:
            QMessageBox.warning(self, "Open Profile", f"{os.path.basename(path)}\n{e.strerror}")
            return
        self.buttons[:] = o.toolbar if "[AppWizard6_ToolbarWizard]" in open(path, encoding="cp1252").read() else []
        self.tbed.fill()

    def make_pic(self, name, size):
        from . import tbicons
        return self.images[name] if name in self.images else tbicons.bmp(name, size)

    def finish(self):
        d = self.designer
        if d is None or d.module.sf is None:
            return
        made = add_toolbar(self.main, d, self.buttons, self.make_pic)
        path = self.save_combo.currentText()
        if path and path != NONE:
            o = G.AppOptions(G.APP_MDI)
            o.toolbar = list(self.buttons)
            text = o.to_rwp()
            if not self.buttons:                # aw7\test1.rwp: only [AppWizard6]
                text = text[:text.index("[AppWizard6_ToolbarWizard]")]
            try:
                with open(path, "w", encoding="cp1252", newline="") as f:
                    f.write(text)
            except OSError as e:
                QMessageBox.warning(self, "Save Profile", f"{path}\n{e.strerror}")
        self.made = made
        self.accept()


def add_toolbar(main, d, buttons, make_pic):
    """the toolbar of the wizard on the form of designer d -> the name of the Toolbar"""
    from lbvc.project.source import parse_source
    m = d.module
    root = d.root
    width = int(root.get("ScaleWidth", root.get("ClientWidth", 4800)) or 4800)
    tb_name = d.new_name("Toolbar")
    il_name = "imlToolbarIcons"
    names = {c.name.lower() for c in d.all_controls()} | d.form_names()
    n = 1
    while il_name.lower() in names:
        n += 1
        il_name = f"imlToolbarIcons{n}"
    o = G.AppOptions(G.APP_MDI)
    o.toolbar = list(buttons)
    f = G.Frm()
    f.depth = 1
    has = any(not b.separator for b in buttons)
    G._write_toolbar(f, o if buttons else _NoButtons(), lambda t: t, name=tb_name,
                     images=il_name if has else None, width=width)
    if has:
        G._write_imagelist(f, il_name, G._image_keys(buttons))
    text = "\r\n".join(["VERSION 5.00", "Begin VB.Form TbWizardTmp "] + f.lines + ["End",
                        'Attribute VB_Name = "TbWizardTmp"']) + "\r\n"
    sf = parse_source(text, None, "Form", True, m.sf.encoding)
    G.attach_pictures(sf, make_pic)
    new = list(sf.root.children)
    d.checkpoint()
    for c in reversed(new):                     # the Toolbar first in the .frm (VB6 writes the newest first)
        d._insert(root, c)
    if c_tab := next((c for c in new if c.name == tb_name), None):
        from .designer import set_prop
        set_prop(c_tab, "TabIndex", d.next_tabindex())
    from . import ocxlib
    lib = ocxlib.LIB_BY_KEY["mscomctllib"]
    if not any(ocxlib.lib_of_object_line(x) is lib for x in m.sf.objects):
        m.sf.objects.append(lib.frm_line)
    p = m.project
    if lib.key not in p.components():
        p.set_components(p.components() + [lib.key])
        if hasattr(main, "refresh_toolbox"):
            main.refresh_toolbox()
    d.build()
    d.changed()
    if has:                                     # Toolbar1_ButtonClick (aw21)
        cases = []
        for b in buttons:
            if not b.separator:
                cases.append(f'        Case "{b.key}"\n' + G._todo(f"the '{b.key}' button", " " * 12))
        code = (f"Private Sub {tb_name}_ButtonClick(ByVal Button As MSComctlLib.Button)\n"
                "    On Error Resume Next\n    Select Case Button.Key\n" + "".join(cases) +
                "    End Select\nEnd Sub\n")
        main._sync_editors()
        old = m.code_text().rstrip("\n")
        new_text = (old + "\n\n" if old.strip() else "") + code
        sw = main.code_windows.get(m)
        if sw is not None:
            sw.widget().editor.setText(new_text)
        m.set_code(new_text)
        m.sync()
        m.dirty = True
    return tb_name


class _NoButtons:
    """a Toolbar without buttons (aw20: Align Top, Height 420, Appearance 1, no ImageList)"""
    toolbar = []


def open_wizard(main):
    d = main.current_designer()
    if d is None:
        QMessageBox.information(main, TITLE, "Open a form in its designer first: the Toolbar Wizard puts the "
                                             "toolbar on the form of the active designer.")
        return None
    dlg = ToolbarWizard(main, d)
    main._tbwizard = dlg                        # the tests reach it
    dlg.exec()
    return dlg.made


__all__ = ["ToolbarWizard", "add_toolbar", "open_wizard"]
