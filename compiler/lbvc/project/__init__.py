# LinBASIC Studio - VB6 project files
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""Readers and writers for VB6 project files.

  vbp.py    .vbp projects and .vbg project groups
  source.py .frm .ctl .dob .pag .cls .bas (designer section, attributes, code)
  frx.py    the binary .frx/.ctx/.dox/.pgx property stores
"""
from .vbp import Project, ProjectGroup, load_project, load_group
from .source import SourceFile, Control, Property, PropertyBlock, load_source
from .frx import FrxStore, FrxWriter
