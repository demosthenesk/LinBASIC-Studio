#!/usr/bin/env python3
# lbvc - the LinBASIC Studio VB6 compiler
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lbvc.driver import main
sys.exit(main())
