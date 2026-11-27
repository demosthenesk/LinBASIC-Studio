"""The Visual Data Manager (Add-Ins > Visual Data Manager..., phase 12.4): a window of its own, as the VisData of VB6
(VB98\\VISDATA.EXE) - the databases of the programs opened through the runtime's own DAO (rtobj), so that what it
shows and changes is what a program sees. Our own code: the VisData sample of VB6 is only read for its behaviour."""


def open_visdata(main):
    """the one Visual Data Manager window of the IDE (shown again when it is open)"""
    from .mainwin import VisDataWindow
    w = getattr(main, "_visdata", None)
    if w is None or not w.isVisible():
        w = VisDataWindow(main)
        main._visdata = w
    w.show()
    w.raise_()
    w.activateWindow()
    return w
