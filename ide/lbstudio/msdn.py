"""Help from the user's own MSDN Library (VB6, October 1998): the .CHM files and their .CHI
companions in the folder set in Tools > Options. Nothing of MSDN is part of LinBASIC Studio.

F1 works as in VB6: the help context number of the type libraries (MsgBox -> 1008978) is looked up
in the #IVB table of the .CHI files, which names the page in the .CHM. Words without a number go
to the keyword index ($WWKeywordLinks); the See Also / Example / Properties links of the pages are
ALinks ($WWAssociativeLinks). The .CHM/.CHI are unpacked once into ~/.cache/linbasic-studio/msdn
with extract_chmLib (package libchm-bin) or 7z (p7zip-full).
"""
import glob
import json
import os
import re
import shutil
import struct
import subprocess

from . import paths

CACHE = paths.cache_dir("msdn")
# the help file a type library names -> the .CHM of MSDN 98 that has it
CHM_OF_HELPFILE = {"vba.hlp": "VBENLR98.CHM", "vb98.chm": "VB98.CHM", "mshflx98.chm": "MSHFLX98.CHM"}


def extractor():
    for tool in ("extract_chmLib", "7z", "7za", "7zz"):
        p = shutil.which(tool)
        if p:
            return tool, p
    if os.name == "nt":
        for d in (os.environ.get("ProgramFiles", ""), os.environ.get("ProgramFiles(x86)", ""),
                  os.environ.get("ProgramW6432", "")):
            p = os.path.join(d, "7-Zip", "7z.exe")
            if d and os.path.exists(p):
                return "7z", p
        p = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "hh.exe")   # HTML Help of Windows
        if os.path.exists(p):
            return "hh", p
    return None, None


def unpack(path, dest):
    """the files of a .CHM or .CHI into dest (once); False when no tool can do it"""
    stamp = os.path.join(dest, ".done")
    if os.path.exists(stamp):
        return True
    tool, exe = extractor()
    if exe is None:
        return False
    os.makedirs(dest, exist_ok=True)
    if tool == "extract_chmLib":
        cmd = [exe, path, dest]
    elif tool == "hh":
        cmd = [exe, "-decompile", dest, path]
    else:
        cmd = [exe, "x", "-y", "-o" + dest, path]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300, check=False, **paths.NO_WINDOW)
    except (OSError, subprocess.SubprocessError):
        return False
    open(stamp, "w").close()
    return True


def _u32(b, o):
    return struct.unpack_from("<I", b, o)[0]


def _read(path):
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError:
        return b""


class ChmFile:
    """one .CHM with its .CHI: help context numbers, keywords, ALinks"""

    def __init__(self, chm, chi):
        self.chm, self.chi = chm, chi
        self.name = os.path.splitext(os.path.basename(chm))[0].lower()
        self._loaded = False
        self.ivb = {}            # context number -> local path (html/vafctMsgBox.htm)
        self.keywords = {}       # keyword lower -> (keyword, [(title, local)])
        self.alinks = {}         # alink name lower -> [(title, local)]

    @property
    def pages_dir(self):
        return os.path.join(CACHE, self.name)

    def load(self):
        if self._loaded:
            return
        self._loaded = True
        src = self.chi or self.chm
        d = os.path.join(CACHE, self.name + ".index")
        if not unpack(src, d):
            return
        strings = _read(os.path.join(d, "#STRINGS"))
        ivb = _read(os.path.join(d, "#IVB"))
        if ivb:
            n = _u32(ivb, 0)
            for i in range(4, min(len(ivb), 4 + n), 8):
                cid, off = struct.unpack_from("<II", ivb, i)
                end = strings.find(b"\0", off)
                self.ivb[cid] = strings[off:end].decode("latin-1")
        topics = _read(os.path.join(d, "#TOPICS"))
        urltbl = _read(os.path.join(d, "#URLTBL"))
        urlstr = _read(os.path.join(d, "#URLSTR"))

        def topic(i):
            o = i * 16
            if o + 16 > len(topics):
                return None
            title_off, url_off = _u32(topics, o + 4), _u32(topics, o + 8)
            title = ""
            if title_off != 0xFFFFFFFF and title_off < len(strings):
                title = strings[title_off:strings.find(b"\0", title_off)].decode("cp1252", "replace")
            if url_off + 12 > len(urltbl):
                return None
            so = _u32(urltbl, url_off + 8) + 8
            local = urlstr[so:urlstr.find(b"\0", so)].decode("latin-1")
            return title, local

        for sub, target in (("$WWKeywordLinks", "keywords"), ("$WWAssociativeLinks", "alinks")):
            bt = _read(os.path.join(d, sub, "BTree"))
            for name, tops in _btree(bt):
                items = [t for t in (topic(i) for i in tops) if t]
                if target == "keywords":
                    old = self.keywords.get(name.lower())
                    self.keywords[name.lower()] = (name, (old[1] if old else []) + items)
                else:
                    self.alinks.setdefault(name.lower(), []).extend(items)

    def page(self, local):
        """the file of a page (the .CHM is unpacked the first time)"""
        if not unpack(self.chm, self.pages_dir):
            return None
        return find_file(self.pages_dir, local)

    def system(self):
        """#SYSTEM of the .CHM: {2: the default topic, 3: the title, 6: the base name}"""
        if getattr(self, "_system", None) is None:
            self._system = {}
            data = b""
            for d in (os.path.join(CACHE, self.name + ".index"), self.pages_dir):
                data = _read(os.path.join(d, "#SYSTEM"))
                if data:
                    break
            p = 4
            while p + 4 <= len(data):
                code, ln = struct.unpack_from("<HH", data, p)
                p += 4
                self._system[code] = data[p:p + ln].split(b"\0")[0].decode("cp1252", "replace")
                p += ln
        return self._system

    @property
    def title(self):
        """"Visual Basic Concepts", "Visual Basic Reference" ... (the Location of the Search tab)"""
        if not unpack(self.chi or self.chm, os.path.join(CACHE, self.name + ".index")):
            return self.name.upper()
        return self.system().get(3) or self.name.upper()

    def toc(self):
        """the table of contents of the .CHM (its .hhc): [TocNode]"""
        if getattr(self, "_toc", None) is None:
            self._toc = []
            if unpack(self.chm, self.pages_dir):
                try:
                    names = [n for n in os.listdir(self.pages_dir) if n.lower().endswith(".hhc")]
                except OSError:
                    names = []
                if names:
                    with open(os.path.join(self.pages_dir, names[0]), "rb") as f:
                        self._toc = parse_sitemap(f.read().decode("cp1252", "replace"), self)
        return self._toc

    def html_files(self):
        """the pages of the .CHM (local names): for the full-text search"""
        if not unpack(self.chm, self.pages_dir):
            return []
        out = []
        for root, _dirs, files in os.walk(self.pages_dir):
            for n in files:
                if n.lower().endswith((".htm", ".html")):
                    out.append(os.path.relpath(os.path.join(root, n), self.pages_dir).replace(os.sep, "/"))
        return sorted(out)


class TocNode:
    """one entry of the Contents tab: a book (children) or a page (chm + local)"""

    def __init__(self, name, chm=None, local=None):
        self.name, self.chm, self.local = name, chm, local
        self.children = []
        self.chm_name = None                    # a page of another .CHM ("vb98.chm::/html/...")

    def __repr__(self):
        return f"TocNode({self.name!r}, {len(self.children)})"


def parse_sitemap(text, chm=None):
    """an HTML Help sitemap (.hhc): <UL> <LI> <OBJECT type="text/sitemap"> <param name="Name" ...> ... as nodes;
    a Local "other.chm::/html/x.htm" names a page of another .CHM of the library"""
    root = TocNode("")
    stack = [root]
    last = None                                 # the entry a <UL> right after it belongs to
    for m in re.finditer(r"<\s*(/?)UL\b[^>]*>|<OBJECT\b[^>]*text/sitemap[^>]*>(.*?)</OBJECT>", text, re.I | re.S):
        if m.group(2) is None:
            if m.group(1):                      # </UL>: back to the level above
                if len(stack) > 1:
                    stack.pop()
            else:                               # <UL>: the children of the last entry (the top list: the root)
                stack.append(last if last is not None else stack[-1])
            last = None
            continue
        params = {k.lower(): v for k, v in
                  re.findall(r"<param\s+name=\"([^\"]+)\"\s+value=\"([^\"]*)\"", m.group(2), re.I)}
        node = TocNode(params.get("name", "").strip(), chm, None)
        local = params.get("local")
        if local:
            mm = re.match(r"^(?:ms-its:|mk:@msitstore:)?(?:.*[\\/])?([^\\/:]+\.chm)::/?(.*)$", local, re.I)
            if mm:
                node.chm_name, node.local = mm.group(1), mm.group(2)
            else:
                node.local = local
        stack[-1].children.append(node)
        last = node
    return root.children


def find_file(root, local):
    """local names are case-insensitive (html/vafctMsgBox.htm is html/vafctmsgbox.htm)"""
    local = local.split("#")[0].lstrip("/")
    p = os.path.join(root, local)
    if os.path.exists(p):
        return p
    cur = root
    for part in local.split("/"):
        try:
            names = {n.lower(): n for n in os.listdir(cur)}
        except OSError:
            return None
        n = names.get(part.lower())
        if n is None:
            return None
        cur = os.path.join(cur, n)
    return cur


def _btree(data):
    """the entries of a $WW*Links/BTree: (name, [topic numbers]) - the listing blocks of 2048 bytes"""
    if len(data) < 0x4C or data[:2] != b";)":
        return
    bs = struct.unpack_from("<H", data, 4)[0] or 2048
    last_listing = _u32(data, 0x1A)
    blk = 0
    seen = set()
    while blk != 0xFFFFFFFF and blk not in seen and blk <= last_listing:
        seen.add(blk)
        base = 0x4C + blk * bs
        if base + 12 > len(data):
            break
        free, count = struct.unpack_from("<HH", data, base)
        nxt = _u32(data, base + 8)
        p = base + 12
        for _ in range(count):
            end = p
            while end + 1 < len(data) and data[end:end + 2] != b"\0\0":
                end += 2
            name = data[p:end].decode("utf-16-le", "replace")
            p = end + 2
            see_also = struct.unpack_from("<H", data, p)[0]
            p += 2 + 2 + 4 + 4
            if see_also:
                p += 4                                   # 1, then the "see also" text
                end = p
                while end + 1 < len(data) and data[end:end + 2] != b"\0\0":
                    end += 2
                p = end + 2
                n = 0
                tops = []
            else:
                n = _u32(data, p)
                p += 4
                tops = list(struct.unpack_from(f"<{n}I", data, p)) if n else []
                p += 4 * n
            p += 8
            yield name, tops
        blk = nxt


class Library:
    """the MSDN folder of the user"""

    def __init__(self, folder):
        self.folder = folder
        self.files = {}
        for chm in sorted(glob.glob(os.path.join(folder, "*"))):
            if chm.lower().endswith(".chm"):
                base = os.path.splitext(chm)[0]
                chi = next((base + e for e in (".CHI", ".chi", ".Chi") if os.path.exists(base + e)), None)
                f = ChmFile(chm, chi)
                self.files[f.name] = f

    @staticmethod
    def available():
        return extractor()[1] is not None

    def file(self, chm_name):
        return self.files.get(os.path.splitext(chm_name.lower())[0])

    def by_context(self, number, helpfile=None):
        """(ChmFile, local) of a help context number"""
        order = []
        if helpfile:
            f = self.file(CHM_OF_HELPFILE.get(helpfile.lower(), helpfile))
            if f:
                order.append(f)
        order += [f for f in self.files.values() if f not in order]
        for f in order:
            f.load()
            if number in f.ivb:
                return f, f.ivb[number]
        return None

    def by_keyword(self, word, main_only=False):
        """[(title, ChmFile, local)] of the keywords that begin with the word"""
        out = []
        w = word.lower().strip()
        names = ["vbenlr98", "vb98"] if main_only else list(self.files)
        for n in names:
            f = self.files.get(n)
            if f is None:
                continue
            f.load()
            for low, (kw, items) in f.keywords.items():
                if low == w or low.startswith(w + " ") or low.startswith(w + ","):
                    for title, local in items:
                        out.append((title or kw, f, local))
        seen, uniq = set(), []
        for t in out:
            key = (t[1].name, t[2].lower())
            if key not in seen:
                seen.add(key)
                uniq.append(t)
        return uniq

    def by_alink(self, names, current=None):
        out = []
        for n in names:
            for f in ([current] if current else []) + [x for x in self.files.values() if x is not current]:
                f.load()
                for title, local in f.alinks.get(n.lower(), []):
                    out.append((title, f, local))
            if out:
                break
        return out

    # ------------------------------------------------------------ Contents (Help > Contents)
    # the Visual Basic part of MSDN 98's table of contents: Visual Basic Documentation is the TOC of VBCON98
    # (Start Page ... Using Visual Basic) and of VBREF98 (Reference, Samples); the books and the data access
    # libraries after it (DEV-8/me3 pictures 2, 3)
    BOOKS = [("Visual Basic Documentation", ["vbcon98", "vbref98"]),
             ("Data Access", ["dao351", "ado200"]),
             ("Books", ["hcvb", "advnvb5"])]
    ROOT = "MSDN Library Visual Studio 6.0"

    def contents(self):
        """[TocNode]: one root, MSDN's books under it, then the .CHMs that no book names, each as its own TOC"""
        if getattr(self, "_contents", None) is None:
            home = self.files.get("vbcon98")
            root = TocNode(self.ROOT, home, home.system().get(2) if home is not None and home.title else None)
            used = set()
            for title, names in self.BOOKS:
                book = TocNode(title)
                for n in names:
                    f = self.files.get(n)
                    if f is not None:
                        used.add(n)
                        book.children += f.toc()
                if book.children:
                    root.children.append(book)
            refd = self._referenced()
            for n, f in self.files.items():
                if n in used or n in refd:
                    continue
                toc = f.toc()
                if len(toc) == 1:
                    node = toc[0]
                elif toc:
                    node = TocNode(f.title)
                    node.children = toc
                else:
                    continue
                same = next((x for x in root.children if x.name == node.name and x.children and node.children), None)
                if same is not None:            # two books of one name (the definitions of VB and of VBA): one
                    have = {c.name.lower() for c in same.children}
                    same.children += [c for c in node.children if c.name.lower() not in have]
                    same.children.sort(key=lambda c: c.name.lower())
                else:
                    root.children.append(node)
            self._contents = [root]
        return self._contents

    def _referenced(self):
        """the .CHMs that the Reference of VBREF98 links to (their pages are in the tree already)"""
        if getattr(self, "_refd", None) is None:
            self._refd = set()
            f = self.files.get("vbref98")
            stack = list(f.toc()) if f is not None else []
            while stack:
                node = stack.pop()
                if node.chm_name:
                    self._refd.add(os.path.splitext(node.chm_name.lower())[0])
                stack.extend(node.children)
        return self._refd

    def resolve(self, node):
        """(ChmFile, local) of a page of the Contents"""
        if node.local is None:
            return None
        f = self.file(node.chm_name) if node.chm_name else node.chm
        return (f, node.local) if f is not None else None

    # ------------------------------------------------------------ Search (Help > Search)
    INDEX_VERSION = 1
    MAX_FOUND = 500                             # VB6's MSDN: "Found: 500" at most

    def fts_path(self):
        return os.path.join(CACHE, "fulltext.json.gz")

    def fts(self, progress=None):
        """the full-text index of the pages of every .CHM (made once, kept in the cache):
        {"pages": [[chm, local, title]], "words": {word: [[page, count]...]}, "titles": {word: [page...]}}"""
        if getattr(self, "_fts", None) is not None:
            return self._fts
        import gzip
        key = sorted(self.files)
        try:
            with gzip.open(self.fts_path(), "rt", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("version") == self.INDEX_VERSION and data.get("files") == key:
                self._fts = data
                return data
        except (OSError, ValueError):
            pass
        pages, words, titles = [], {}, {}
        files = list(self.files.values())
        for i, f in enumerate(files):
            if progress is not None and progress(i, len(files), f.name) is False:
                return None
            for local in f.html_files():
                path = find_file(f.pages_dir, local)
                try:
                    with open(path, "rb") as fh:
                        raw = fh.read()
                except (OSError, TypeError):
                    continue
                html = raw.decode("cp1252", "replace")
                tm = re.search(r"<title>(.*?)</title>", html, re.I | re.S)
                title = re.sub(r"\s+", " ", tm.group(1)).strip() if tm else os.path.basename(local)
                pid = len(pages)
                pages.append([f.name, local, title])
                counts = {}
                for w in WORD_RE.findall(page_text(html).lower()):
                    counts[w] = counts.get(w, 0) + 1
                for w, c in counts.items():
                    words.setdefault(w, []).append([pid, c])
                for w in set(WORD_RE.findall(title.lower())):
                    titles.setdefault(w, []).append(pid)
        data = {"version": self.INDEX_VERSION, "files": key, "pages": pages, "words": words, "titles": titles}
        try:
            os.makedirs(CACHE, exist_ok=True)
            with gzip.open(self.fts_path(), "wt", encoding="utf-8") as f:
                json.dump(data, f)
        except OSError:
            pass
        self._fts = data
        return data

    def search(self, query, titles_only=False, similar=True, within=None, progress=None):
        """Search: [(title, ChmFile, local, rank)], best first, at most 500. The words of the query must all be
        in the page (AND); OR / NOT / NEAR between words as MSDN's operators (NEAR as AND); "a phrase" as its
        words; similar: words that begin with the word (msgbox -> msgboxes); within: the pages of the previous
        results"""
        data = self.fts(progress)
        if data is None:
            return []
        table = data["titles"] if titles_only else data["words"]
        vocab = list(table)

        def hits(word):
            out = {}
            ws = [word]
            if similar:
                ws = [v for v in vocab if v == word or (v.startswith(word) and len(v) - len(word) <= 3)]
            for v in ws:
                for x in table.get(v, []):
                    pid, c = (x, 1) if titles_only else (x[0], x[1])
                    out[pid] = out.get(pid, 0) + c
            return out

        tokens = re.findall(r'"[^"]*"|\S+', query)
        result, op = None, "AND"
        for tok in tokens:
            up = tok.upper()
            if up in ("AND", "OR", "NOT", "NEAR"):
                op = up
                continue
            for w in WORD_RE.findall(tok.strip('"').lower()):
                h = hits(w)
                if result is None:
                    result = dict(h) if op != "NOT" else {}
                elif op == "OR":
                    for k, v in h.items():
                        result[k] = result.get(k, 0) + v
                elif op == "NOT":
                    result = {k: v for k, v in result.items() if k not in h}
                else:                           # AND, NEAR
                    result = {k: v + h[k] for k, v in result.items() if k in h}
                op = "AND"
        if not result:
            return []
        if not titles_only:                     # the words in the title of a page count more (MsgBox Function first)
            words = [w for tok in tokens if tok.upper() not in ("AND", "OR", "NOT", "NEAR")
                     for w in WORD_RE.findall(tok.strip('"').lower())]
            titles = data["titles"]
            for w in words:
                for pid in titles.get(w, []):
                    if pid in result:
                        result[pid] += 50
        if within is not None:
            keep = {(c.name, loc) for _t, c, loc, _r in within}
            pages = data["pages"]
            result = {k: v for k, v in result.items() if (pages[k][0], pages[k][1]) in keep}
        best = sorted(result.items(), key=lambda kv: (-kv[1], data["pages"][kv[0]][2].lower()))[:self.MAX_FOUND]
        out = []
        for rank, (pid, _score) in enumerate(best, 1):
            chm, local, title = data["pages"][pid]
            f = self.files.get(chm)
            if f is not None:
                out.append((title, f, local, rank))
        return out

    def all_keywords(self):
        out = {}
        for f in self.files.values():
            f.load()
            for low, (kw, items) in f.keywords.items():
                out.setdefault(low, kw)
        return sorted(out.values(), key=str.lower)


WORD_RE = re.compile(r"[a-z0-9_$#]+", re.I)


def page_text(html):
    """the words of a page: without the scripts, the styles and the tags"""
    html = re.sub(r"<(script|style)\b.*?</\1>", " ", html, flags=re.I | re.S)
    html = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"&[a-z]+;|&#\d+;", " ", html, flags=re.I)


# the help file of each type library (index.json of data/vb6/typelibs)
def helpfile_of_lib(lib):
    try:
        with open(os.path.join(paths.VB6_DATA, "typelibs", "index.json"), encoding="utf-8") as f:
            for x in json.load(f):
                if x["short"] == lib:
                    return x.get("helpfile") or None
    except OSError:
        pass
    return None


# ---------------------------------------------------------------- the pages
ALINK_RE = re.compile(r"<object[^>]*>(?:(?!</object>).)*?<param\s+name=\"Command\"\s+value=\"([AK]Link)[^\"]*\">"
                      r"(?:(?!</object>).)*?</object>", re.I | re.S)


def rewrite(html):
    """the HHCtrl objects of the pages as links that the viewer follows: the See Also / Example
    buttons (an object with a Text), and the words that open a definition (an object whose link is
    the <a href=JavaScript:alink_4.Click()> after it)"""
    targets = {}

    def repl(m):
        block = m.group(0)
        kind = m.group(1).lower()
        oid = re.search(r"\bid=\"?([A-Za-z0-9_]+)", block)
        text = re.search(r"name=\"Text\"\s+value=\"(?:Text:)?([^\"]*)\"", block, re.I)
        item2 = re.search(r"name=\"Item2\"\s+value=\"([^\"]*)\"", block, re.I)
        target = f"{kind}:{item2.group(1) if item2 else ''}"
        if oid:
            targets[oid.group(1).lower()] = target
        if text and text.group(1).strip():
            return f'<a href="{target}">{text.group(1)}</a>'
        return ""
    html = ALINK_RE.sub(repl, html)
    html = re.sub(r"<object[^>]*>.*?</object>", "", html, flags=re.I | re.S)
    html = re.sub(r"href=\"?JavaScript:([A-Za-z0-9_]+)\.Click\(\)\"?",
                  lambda m: f'href="{targets.get(m.group(1).lower(), "#")}"', html, flags=re.I)
    return html
