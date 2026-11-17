# LinBASIC Studio - tests for the type library reader
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
# Needs your own VB6 installation, extracted with tools/rea/extract-vb6-msdn.sh:
#   VB6_REA=<output folder>/VB6-REA python3 -m unittest discover -s tests/vb6
import os, sys, unittest
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "compiler"))
from lbvc.typelib import TypeLib

REA = os.environ.get("VB6_REA", "")


@unittest.skipUnless(os.path.isdir(REA), "set VB6_REA to the VB6-REA folder")
class TestTypeLib(unittest.TestCase):
    def test_vb_library(self):
        lib = TypeLib.load(os.path.join(REA, "VB98", "VB6.OLB"))
        self.assertEqual((lib.name, lib.guid), ("VB", "{FCFB3D2E-A0FA-1068-A738-08002B3371B5}"))
        cb = lib.by_name["commandbutton"]
        self.assertEqual((cb.default_interface, cb.event_interface), ("_CommandButton", "CommandButtonEvents"))
        ev = [m.name for m in lib.by_name["commandbuttonevents"].members]
        self.assertIn("Click", ev)
        cap = [m for m in lib.by_name["_commandbutton"].members if m.name == "Caption"]
        self.assertEqual(sorted(m.invkind for m in cap), ["get", "let"])

    def test_vba_runtime(self):
        lib = TypeLib.load(os.path.join(REA, "System32", "msvbvm60.dll"), 1)
        self.assertEqual(lib.name, "VBA")
        consts = {m.name: m.value for m in lib.by_name["vbmsgboxstyle"].members}
        self.assertEqual((consts["vbYesNo"], consts["vbQuestion"]), (4, 32))
        instr = [m for m in lib.by_name["strings"].members if m.name == "InStr"][0]
        self.assertEqual([p.name for p in instr.vb_params()], ["Start", "String1", "String2", "Compare"])
        self.assertEqual(instr.vb_params()[3].default, 0)

    def test_vbrun_colors(self):
        lib = TypeLib.load(os.path.join(REA, "System32", "msvbvm60.dll"), 3)
        c = {m.name: m.value for m in lib.by_name["colorconstants"].members}
        self.assertEqual((c["vbRed"], c["vbBlue"], c["vbWhite"]), (255, 16711680, 16777215))


if __name__ == "__main__":
    unittest.main()
