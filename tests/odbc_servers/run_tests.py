#!/usr/bin/env python3
# LinBASIC Studio - ADO over ODBC against real database servers (MariaDB / MySQL, PostgreSQL): a check
# by hand, without a VB6 oracle (VB6 would need the same servers). The servers must be running:
# MariaDB on 127.0.0.1:3306 with a database lbtest for lb/lb, PostgreSQL on 127.0.0.1:5432 with a
# database lbtest for postgres (trust). The output is in work/srv/expected/srv.txt.
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
import os, shutil, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', 'compiler'))
from lbvc.driver import compile_project, build_exe, find_runtime   # noqa: E402

work = os.path.join(HERE, 'work', 'srv')
shutil.rmtree(work, ignore_errors=True)
os.makedirs(os.path.join(work, 'expected'))
csrc, _ = compile_project([os.path.join(HERE, 'srv', 'SrvTest.vbp')])
exe = build_exe(csrc, os.path.join(work, 'prog'), find_runtime()) or os.path.join(work, 'prog')
r = subprocess.run([exe], cwd=work, capture_output=True, text=True, timeout=120,
                   env=dict(os.environ, QT_QPA_PLATFORM='offscreen', VBRT_NO_ERRBOX='1'))
print(open(os.path.join(work, 'expected', 'srv.txt'), encoding='cp1252', errors='replace').read())
if r.returncode: print('[exit', r.returncode, ']', r.stderr)
