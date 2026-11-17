"""File > Print (phase 8.3): the code, the form as text and the form image, as the VB6 IDE prints them.

What the VB6 IDE printed (VB-TESTS/DEV-8/pr1, through the Microsoft XPS Document Writer):
  - Courier New 10, from the top left corner of the printable area (no margins of its own)
  - every module on new pages; a page begins with "<module> - <page>" (the pages of each module counted from 1)
    and an empty line; a page is full when the next line would pass its bottom (A4: 79 lines with these two)
  - a form: its code, then Form As Text (its .frm before the Attribute lines, on new pages, counted from 1 again),
    then Form Image (the client area as on the screen, 1 pixel = 1/96 inch, no title)
  - Current Project: the modules in the order of the project; Selection: the selected text, with the header
  - lines: 10 pt (em) at the printer's resolution, the line pitch 88/83 of it, the first base line 65/83 of it
    (the numbers of the XPS at 600 dpi: em 83, pitch 88, base line 65)
"""

PITCH = 88 / 83
BASE = 65 / 83


def text_pages(name, lines, per_page):
    """the pages of one text: [[header, "", line, ...], ...]"""
    lines = list(lines)
    while lines and not lines[-1].strip():
        lines.pop()                             # (the empty lines at the end of a module are not printed)
    body = max(1, per_page - 2)
    pages = []
    for k in range(0, max(1, len(lines)), body):
        pages.append([f"{name} - {len(pages) + 1}", ""] + lines[k:k + body])
    return pages


def form_text_lines(module):
    """Form As Text: the .frm before the Attribute lines"""
    sf = getattr(module, "sf", None)
    if sf is None:
        return []
    lines = list(sf.lines())[:module.code_offset()]
    while lines and lines[-1].startswith("Attribute "):
        lines.pop()
    return lines


def jobs(modules, code_of, what, selection=None):
    """what to print: [("text", name, lines) | ("image", module)]; modules in the order to print; what: set of
    "code", "text", "image"; selection: (module, text) for Range Selection"""
    out = []
    if selection is not None:
        m, text = selection
        return [("text", m.name, text.replace("\r\n", "\n").split("\n"))]
    for m in modules:
        form = m.kind in ("Form", "UserControl", "PropertyPage", "UserDocument") and getattr(m, "sf", None) is not None
        if "code" in what:
            code = code_of(m).replace("\r\n", "\n")
            if code.strip():
                out.append(("text", m.name, code.split("\n")))
        if form and "text" in what:
            out.append(("text", m.name, form_text_lines(m)))
        if form and "image" in what:
            out.append(("image", m))
    return out


def layout(job_list, per_page):
    """the pages: ("text", [lines]) or ("image", module)"""
    pages = []
    for j in job_list:
        if j[0] == "text":
            pages += [("text", p) for p in text_pages(j[1], j[2], per_page)]
        else:
            pages.append(("image", j[1]))
    return pages


def metrics(dpi):
    em = round(10 * dpi / 72)
    return em, round(em * PITCH), round(em * BASE)


def print_pages(printer, pages, image_of):
    """the pages on the printer (QPrinter); image_of(module) -> QPixmap of the form (screen pixels). The number
    of pages printed"""
    from PyQt6.QtGui import QFont, QPainter
    dpi = printer.resolution()
    em, pitch, base = metrics(dpi)
    p = QPainter()
    if not p.begin(printer):
        return 0
    try:
        f = QFont("Courier New")
        f.setStyleHint(QFont.StyleHint.TypeWriter)
        f.setPixelSize(em)
        p.setFont(f)
        for i, (kind, data) in enumerate(pages):
            if i:
                printer.newPage()
            if kind == "text":
                for k, line in enumerate(data):
                    if line:
                        p.drawText(0, k * pitch + base, line.expandtabs(4))
            else:
                pm = image_of(data)
                if pm is not None and not pm.isNull():
                    s = dpi / 96
                    p.drawPixmap(0, 0, round(pm.width() * s), round(pm.height() * s), pm)
    finally:
        p.end()
    return len(pages)


def lines_per_page(printer):
    dpi = printer.resolution()
    _, pitch, _ = metrics(dpi)
    h = printer.pageLayout().paintRectPixels(dpi).height()
    return max(3, h // pitch)
