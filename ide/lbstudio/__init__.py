"""LinBASIC Studio — the IDE.

A development environment for projects written in the VB6 dialect (.vbp .frm .bas .cls),
built on the editor of Linpad++ (editor.py, styles.py, langs.py, settings.py and the
Scintilla data files are taken from Linpad++, which stays an independent project).
The compiler is lbvc (compiler/), the runtime libvbrt (runtime/).

Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""

import os as _os
import sys as _sys


def _qt5_as_qt6():
    """the IDE is written for PyQt6; where there is only PyQt5 (Windows 7, 32-bit Windows: the IDE of
    Windows bundles Python 3.8 and PyQt5 5.15) PyQt5 is used under the names of PyQt6, with the few
    differences of Qt 6 made good here. LBSTUDIO_QT=5 asks for it where both are."""
    import PyQt5
    from PyQt5 import QtCore, QtGui, QtWidgets
    _sys.modules["PyQt6"] = PyQt5
    for name in ("QtCore", "QtGui", "QtWidgets", "QtNetwork", "QtTest", "Qsci", "QtPrintSupport"):
        try:
            mod = __import__("PyQt5." + name, fromlist=["x"])
        except ImportError:
            continue
        _sys.modules["PyQt6." + name] = mod
        setattr(PyQt5, name, mod)
    # classes of QtWidgets in Qt 5 that Qt 6 has in QtGui
    for n in ("QAction", "QActionGroup", "QShortcut", "QUndoStack", "QUndoCommand", "QFileSystemModel"):
        if hasattr(QtWidgets, n) and not hasattr(QtGui, n):
            setattr(QtGui, n, getattr(QtWidgets, n))
    # the positions of the events: position() / globalPosition() of Qt 6
    QtGui.QMouseEvent.position = lambda self: self.localPos()
    QtGui.QMouseEvent.globalPosition = lambda self: self.screenPos()
    QtGui.QDropEvent.position = lambda self: self.posF()
    QtGui.QHoverEvent.position = lambda self: self.posF()
    QtGui.QEnterEvent.position = lambda self: self.localPos()
    QtGui.QEnterEvent.globalPosition = lambda self: self.screenPos()
    # QFontDatabase: static in Qt 6
    _fdb = QtGui.QFontDatabase

    class QFontDatabase(_fdb):
        @staticmethod
        def families(*a):
            return _fdb().families(*a)

        @staticmethod
        def styles(family):
            return _fdb().styles(family)

        @staticmethod
        def isFixedPitch(family, style=""):
            return _fdb().isFixedPitch(family, style)

        @staticmethod
        def pointSizes(family, style=""):
            return _fdb().pointSizes(family, style)
    QtGui.QFontDatabase = QFontDatabase
    QtCore.QT_VERSION_STR6 = False
    QtWidgets.QApplication.setAttribute(QtCore.Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)


if "PyQt6" not in _sys.modules:
    try:
        if _os.environ.get("LBSTUDIO_QT") == "5":
            raise ImportError
        import PyQt6.QtCore  # noqa: F401
    except ImportError:
        _qt5_as_qt6()

APP_NAME = "LinBASIC Studio"
VERSION = "1.1.2"
