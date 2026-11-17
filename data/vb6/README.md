# data/vb6

The VB6 object model and IDE command set, recovered in Phase 0 (REA) from a local Visual
Basic 6.0 installation. The compiler, the runtime and the IDE (IntelliSense, Object Browser,
Properties window, F1 help) are built from these files.

| File | Content | Made by |
|---|---|---|
| `typelibs/*.json` | one file per type library: types, members, signatures, DISPIDs, enum values, helpcontext ids; for VBA the MSVBVM60 export of each runtime function | `tools/rea/dump_typelibs.py` |
| `typelibs/index.json` | the libraries, with GUID, version and help file | `tools/rea/dump_typelibs.py` |
| `ide_commands.json` | IDE menus and commands, dialog titles, keyboard shortcuts | `tools/rea/dump_resources.py` |
| `runtime_errors.json` | trappable run-time errors and their messages (`Err.Description`) | `tools/rea/dump_resources.py` |

No Microsoft binary or help text is stored here: help strings are left out on purpose, and the
IDE opens the user's own MSDN help through each member's `helpcontext`.

Regenerate from your own VB6 installation (`sudo tools/rea/extract-vb6-msdn.sh <disk.vdi> <folder>`
copies it read-only from the disk of a Windows virtual machine):

    python3 tools/rea/dump_typelibs.py <folder>/VB6-REA data/vb6/typelibs
    python3 tools/rea/dump_resources.py <folder>/VB6-REA data/vb6
    python3 tools/rea/make_feature_map.py
