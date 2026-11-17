# LinBASIC Studio - lbvc, the VB6 compiler
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""lbvc compiles VB6 projects (.vbp, .frm, .bas, .cls, .ctl) to native programs
through C (tcc/gcc), linked with the libvbrt runtime.

  project/   readers and writers for the VB6 project files (.vbp .vbg .frm .frx .cls .bas ...)
  typelib.py COM type libraries (.tlb .olb and TYPELIB resources of .dll/.ocx)
"""
__version__ = "0.1.0"
