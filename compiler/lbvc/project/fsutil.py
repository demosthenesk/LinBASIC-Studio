# LinBASIC Studio - file names of VB6 projects on Linux
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""VB6 projects come from Windows, where file names do not care about case: a .vbp says
Form=main.frm for MAIN.FRM, a form says "BUTTON.frx" for BUTTON.FRX.  On Linux the name is
found again by its letters."""
import os


def ci_path(path):
    """path itself when it exists, else the existing file whose name differs only in case;
    the path as given when there is none"""
    if not path or os.path.exists(path):
        return path
    path = os.path.normpath(path)
    head, parts = path, []
    while head and not os.path.exists(head):
        head, tail = os.path.split(head)
        if not tail:
            break
        parts.insert(0, tail)
    cur = head or "."
    for part in parts:
        try:
            names = os.listdir(cur)
        except OSError:
            return path
        hit = next((n for n in names if n == part), None) or next((n for n in names if n.lower() == part.lower()), None)
        if hit is None:
            return path
        cur = os.path.join(cur, hit)
    return cur


def _short_match(short, name):
    """name fits the MS-DOS short name short (FORM1~1.FRM ~ Form1 Main.frm)"""
    sb, _, se = short.upper().partition('.')
    nb, _, ne = name.upper().partition('.') if '.' in name else (name.upper(), '', '')
    nb = name.upper().rsplit('.', 1)[0].replace(' ', '').replace('.', '')
    ne = name.upper().rsplit('.', 1)[1] if '.' in name else ''
    if '~' not in sb:
        return False
    pre = sb.split('~')[0]
    return nb.startswith(pre) and ne[:3] == se[:3]


def find_moved(base, rel, depth=4):
    """a member of a project whose path is not where the project says (an absolute path of the
    author's computer, short MS-DOS names): the one file of that name in the project's folders"""
    name = rel.replace("\\", "/").rstrip("/").split("/")[-1]
    if not name:
        return None
    hits, short = [], []
    root = os.path.normpath(base)
    for d, dirs, files in os.walk(root):
        if d[len(root):].count(os.sep) >= depth:
            dirs[:] = []
        for f in files:
            if f.lower() == name.lower():
                hits.append(os.path.join(d, f))
            elif '~' in name and _short_match(name, f):
                short.append(os.path.join(d, f))
    if len(hits) >= 1:
        return sorted(hits, key=len)[0]
    if len(short) == 1:
        return short[0]
    return None
