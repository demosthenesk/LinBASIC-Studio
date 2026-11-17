"""Start LinBASIC Studio:  python3 -m lbstudio [project.vbp]"""
import argparse
import os
import sys

from . import APP_NAME, VERSION, paths, settings


def main(argv=None):
    ap = argparse.ArgumentParser(prog="lbstudio", description=f"{APP_NAME} {VERSION}")
    ap.add_argument("project", nargs="?", help="a .vbp project to open")
    ap.add_argument("--settingsDir", help="keep the settings in this folder")
    a = ap.parse_args(argv)
    if a.settingsDir:
        paths.set_config_dir(a.settingsDir)
    paths.ensure_user_files()
    settings.init()

    from PyQt6.QtWidgets import QApplication
    from .mainwindow import MainWindow
    app = QApplication(sys.argv[:1])
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(VERSION)
    app.setDesktopFileName("linbasic-studio")
    from PyQt6.QtGui import QIcon
    icon = os.path.join(paths.RES_DIR, "linbasic-studio-256.png")   # no SVG plugin needed
    if os.path.exists(icon):
        app.setWindowIcon(QIcon(icon))
    w = MainWindow()
    w.restore_layout()
    if a.project and os.path.exists(a.project):
        w.open_project(a.project)
    w.show()
    return app.exec()
