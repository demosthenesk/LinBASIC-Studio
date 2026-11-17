# LinBASIC Studio - tests for the VB6 project-file readers and writers
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 -m unittest discover -s tests/vb6       (from the source folder)
import os, sys, unittest, tempfile
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "compiler"))
from lbvc.project import load_project, load_group, FrxStore
from lbvc.project import load_source as _load_source
load_source = lambda path: _load_source(path, encoding="cp1253")
from lbvc.project.source import FrxRef, parse_source, Control
from lbvc.project.frx import FrxWriter

S = os.path.join(HERE, "samples")
p = lambda n: os.path.join(S, n)


def read(n):
    with open(p(n), "rb") as f:
        return f.read()


class TestVbp(unittest.TestCase):
    def test_fields(self):
        pr = load_project(p("Project1.vbp"))
        self.assertEqual(pr.name, "Project1")
        self.assertEqual(pr.type, "Exe")
        self.assertEqual(pr.startup, "Form1")
        self.assertEqual(pr.exe_name, "Demo.exe")
        self.assertEqual(list(pr.members()), [("Form", "Form1", "Form1.frm"),
                                              ("Module", "Module1", "Module1.bas"),
                                              ("Class", "CCounter", "CCounter.cls")])
        ref = pr.references()[0]
        self.assertEqual((ref.kind, ref.guid, ref.version, ref.description),
                         ("G", "{00020430-0000-0000-C000-000000000046}", "2.0", "OLE Automation"))
        self.assertEqual(pr.objects()[0].file, "MSCOMCTL.OCX")
        self.assertEqual(pr.sections[0][0], "MS Transaction Server")

    def test_roundtrip(self):
        pr = load_project(p("Project1.vbp"))
        self.assertEqual(pr.dumps().encode("cp1253"), read("Project1.vbp"))

    def test_edit(self):
        pr = load_project(p("Project1.vbp"))
        pr.set("ExeName32", "New.exe", quote=True)
        self.assertIn('ExeName32="New.exe"\r\n', pr.dumps())

    def test_group(self):
        g = load_group(p("Group1.vbg"))
        self.assertEqual(g.startup, "Project1.vbp")
        self.assertEqual(g.projects(), ["Project1.vbp", "Lib\\Lib.vbp"])
        self.assertEqual(g.dumps().encode(), read("Group1.vbg"))


class TestForm(unittest.TestCase):
    def setUp(self):
        self.f = load_source(p("Form1.frm"))

    def test_roundtrip_bytes(self):
        for n in ("Form1.frm", "CCounter.cls", "Module1.bas"):
            self.assertEqual(load_source(p(n)).dumps().encode("cp1253"), read(n), n)

    def test_tree(self):
        f = self.f
        self.assertEqual(f.kind, "Form")
        self.assertEqual(f.name, "Form1")
        self.assertEqual(f.root.get("Caption"), 'Demo "quoted"')
        self.assertEqual(f.root.block("Font").get("Size"), 8.25)
        names = [(c.type_name, c.name, c.index) for c in f.controls()]
        self.assertEqual(names[:5], [("Form", "Form1", None), ("ListView", "ListView1", None),
                                     ("Frame", "Frame1", None), ("OptionButton", "optChoice", 1),
                                     ("OptionButton", "optChoice", 0)])
        frame = f.root.children[1]
        self.assertTrue(all(c.parent is frame for c in frame.children))
        self.assertEqual(frame.children[1].get("Value"), -1)
        menu = f.root.children[-1]
        self.assertEqual(menu.children[0].get("Shortcut"), "^Q")
        lv = f.root.children[0]
        self.assertEqual(lv.library, "MSComctlLib")
        col2 = lv.block("ColumnHeader(2)")
        self.assertEqual(col2.guid, "{BDD1F052-858B-11D1-B16A-00C0F0283628}")
        self.assertEqual(col2.get("Text"), "Value")
        self.assertEqual(lv.get("BackColor"), -2147483643)
        self.assertEqual(f.root.prop("StartUpPosition").comment, "Windows Default")

    def test_attributes_and_code(self):
        f = self.f
        self.assertEqual(f.attribute("VB_PredeclaredId"), "True")
        self.assertEqual(f.code[0], "Option Explicit")
        self.assertIn('Debug.Print "Γεια σου"', f.code_text)

    def test_edit_writes_vb6_layout(self):
        f = self.f
        f.root.set("Caption", "Νέο")
        btn = Control("VB.CommandButton", "cmdOK")
        btn.set("Caption", "OK")
        btn.set("Default", -1)
        f.root.add_child(btn)
        out = f.dumps()
        self.assertIn('   Caption         =   "Νέο"\r\n', out)
        self.assertIn("   Begin VB.CommandButton cmdOK \r\n"
                      '      Caption         =   "OK"\r\n'
                      "      Default         =   -1\r\n"
                      "   End\r\nEnd\r\nAttribute VB_Name", out)
        g = parse_source(out.replace("\r\n", "\n"))
        self.assertEqual(g.root.children[-1].name, "cmdOK")
        # untouched lines are kept exactly
        self.assertIn("      Object.Width           =   2540\r\n", out)

    def test_frx(self):
        refs = self.f.frx_refs()
        self.assertEqual(sorted(r.name for _, r in refs), ["Icon", "ItemData", "List", "Text"])
        store = FrxStore.load(p("Form1.frx"), [r.value.offset for _, r in refs], encoding="cp1253")
        d = {r.name: store.decode(r.name, r.value) for _, r in refs}
        self.assertEqual(d["Icon"][0], "ico")
        self.assertEqual(len(d["Icon"][1]), 24)
        self.assertEqual(d["List"], ["Alpha", "Beta"])
        self.assertEqual(d["ItemData"], [10, 20])
        self.assertEqual(d["Text"], "Line 1\r\nΓραμμή 2")


class TestClassAndModule(unittest.TestCase):
    def test_class(self):
        c = load_source(p("CCounter.cls"))
        self.assertEqual(c.kind, "Class")
        self.assertEqual(c.name, "CCounter")
        self.assertEqual([(x.name, x.value) for x in c.class_props][:2],
                         [("MultiUse", -1), ("Persistable", 0)])
        self.assertIn("Attribute Value.VB_UserMemId = 0", c.code)   # procedure attribute stays in code

    def test_module(self):
        m = load_source(p("Module1.bas"))
        self.assertEqual((m.kind, m.name), ("Module", "Module1"))
        self.assertEqual(m.code[-2:], ["Public Sub Main()", "End Sub"])


class TestFrxWriter(unittest.TestCase):
    def test_vb6_written_lists(self):
        # bytes of a ListBox/ComboBox saved by VB6 itself (a test form saved in the VB6 IDE)
        frx = bytes.fromhex("0300030002003130020032300200 2d31"
                            "030005000500416c70686104004265746105004761 6d6d61"
                            "02000500040031303030050037303030 30"
                            "0200030003004f6e650300 54776f".replace(" ", ""))
        s = FrxStore(frx, [0x00, 0x10, 0x28, 0x39])
        self.assertEqual(s.itemdata(0x00), [10, 20, -1])
        self.assertEqual(s.list(0x10), ["Alpha", "Beta", "Gamma"])
        self.assertEqual(s.itemdata(0x28), [1000, 70000])
        self.assertEqual(s.list(0x39), ["One", "Two"])
        w = FrxWriter()
        w.itemdata([1000, 70000])
        w.list(["One", "Two"])
        self.assertEqual(w.bytes(), frx[0x28:])          # same bytes as VB6

    def test_empty_list(self):
        self.assertEqual(FrxStore(b"\0\0\0\0", [0, 2]).list(0), [])
        w = FrxWriter()
        w.list([])
        self.assertEqual(w.bytes(), b"\0\0")

    def test_roundtrip(self):
        w = FrxWriter()
        a = w.picture(b"BM" + b"\0" * 30)
        b = w.list(["one", "two", "three"])
        c, _ = w.string("multi\r\nline")
        d = w.itemdata([1, -2])
        s = FrxStore(w.bytes(), [a, b, c, d])
        self.assertEqual(s.picture(a)[0], "bmp")
        self.assertEqual(s.list(b), ["one", "two", "three"])
        self.assertEqual(s.decode("Text", FrxRef("x.frx", c, True)), "multi\r\nline")
        self.assertEqual(s.itemdata(d), [1, -2])


if __name__ == "__main__":
    unittest.main()
