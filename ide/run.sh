#!/bin/sh
# Run LinBASIC Studio from the source tree (needs PyQt6 and PyQt6-QScintilla).
HERE="$(dirname "$(readlink -f "$0")")"
PYTHONPATH="$HERE${PYTHONPATH:+:$PYTHONPATH}" exec python3 -m lbstudio "$@"
