# VB6 feature map (Phase 0, REA)

Generated from a local Visual Basic 6.0 installation (VB6.EXE 6.0 SP6, MSVBVM60.DLL, OCX files, MSDN Library October 1998) by `tools/rea/dump_typelibs.py`, `tools/rea/dump_resources.py` and `tools/rea/make_feature_map.py`. Data: `data/vb6/`.

## Type libraries

| Library | File | Version | Types | Members | Help file |
|---|---|---|---|---|---|
| stdole | stdole2.tlb | 2.0 | 42 | 89 | — |
| VBA | msvbvm60.dll | 6.0 | 29 | 371 | VBA.HLP |
| VBRUN | msvbvm60.dll | 6.0 | 96 | 735 | VB98.CHM |
| VB | VB6.OLB | 6.0 | 98 | 3132 | VB98.CHM |
| VBInternal | VB6.EXE | 6.0 | 20 | 61 | z:\__VB_HELP__ |
| VBIDE | VB6EXT.OLB | 5.3 | 112 | 669 | VBEXT98.CHM |
| MSComctlLib | mscomctl.OCX | 2.2 | 132 | 1113 | cmctl198.chm |
| MSComCtl2 | mscomct2.ocx | 2.0 | 42 | 420 | cmctl298.chm |
| ComCtl3 | ComCt332.ocx | 1.1 | 20 | 135 | cmctl398.chm |
| MSComDlg | ComDlg32.OCX | 1.2 | 12 | 171 | cmdlg98.chm |
| MSFlexGridLib | MSFlxGrd.ocx | 1.0 | 30 | 334 | MSHFlx98.chm |
| RichTextLib | richtx32.Ocx | 1.2 | 28 | 226 | RTFBox98.CHM |
| TabDlg | TabCtl32.Ocx | 1.1 | 16 | 134 | mstab98.chm |
| Scripting | scrrun.dll | 1.0 | 32 | 149 | VBENLR98.CHM |
| ADODB | msado28.tlb | 2.8 | 97 | 734 | ado270.chm |

## Intrinsic controls and objects (library VB)

| Object | Properties | Methods | Events |
|---|---|---|---|
| PictureBox | 66 | 22 | 26: Click, Change, DblClick, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LinkClose, LinkError, LinkOpen … |
| Label | 46 | 10 | 18: Click, Change, DblClick, DragDrop, DragOver, LinkClose, LinkError, LinkOpen, MouseDown, MouseMove, MouseUp, LinkNotify … |
| TextBox | 56 | 11 | 24: Change, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LinkClose, LinkError, LinkOpen, LostFocus, LinkNotify … |
| Frame | 35 | 6 | 13: DragDrop, DragOver, MouseDown, MouseMove, MouseUp, Click, DblClick, OLEDragOver, OLEDragDrop, OLEGiveFeedback, OLEStartDrag, OLESetData … |
| CommandButton | 43 | 7 | 17: Click, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove, MouseUp, OLEDragOver … |
| CheckBox | 48 | 7 | 18: Click, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove, MouseUp, OLEDragOver … |
| OptionButton | 43 | 7 | 19: Click, DblClick, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove, MouseUp … |
| ComboBox | 54 | 10 | 19: Change, Click, DblClick, DragDrop, DragOver, DropDown, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, OLEDragOver … |
| ListBox | 54 | 10 | 21: Click, DblClick, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove, MouseUp … |
| HScrollBar | 27 | 6 | 10: Change, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, Scroll, Validate |
| VScrollBar | 27 | 6 | 10: Change, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, Scroll, Validate |
| Timer | 6 | 0 | 1: Timer |
| Printer | 42 | 12 | 0 |
| Form | 64 | 22 | 31: Load, DragDrop, DragOver, LinkClose, LinkError, LinkExecute, LinkOpen, Resize, Unload, QueryUnload, Activate, Deactivate … |
| Screen | 10 | 0 | 0 |
| Clipboard | 0 | 6 | 0 |
| DriveListBox | 38 | 7 | 16: Change, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, OLEDragOver, OLEDragDrop, OLEGiveFeedback, OLEStartDrag … |
| DirListBox | 39 | 7 | 20: Change, Click, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove, MouseUp … |
| FileListBox | 48 | 7 | 22: Click, DblClick, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove, MouseUp … |
| Menu | 10 | 0 | 1: Click |
| MDIForm | 33 | 10 | 25: Load, DragDrop, DragOver, LinkClose, LinkError, LinkExecute, LinkOpen, Resize, Unload, QueryUnload, Activate, Deactivate … |
| App | 30 | 2 | 0 |
| Shape | 19 | 3 | 0 |
| Line | 14 | 2 | 0 |
| Image | 28 | 6 | 13: Click, DblClick, DragDrop, DragOver, MouseDown, MouseMove, MouseUp, OLEDragOver, OLEDragDrop, OLEGiveFeedback, OLEStartDrag, OLESetData … |
| Data | 44 | 8 | 15: Validate, Error, Reposition, DragDrop, DragOver, MouseDown, MouseMove, MouseUp, Resize, OLEDragOver, OLEDragDrop, OLEGiveFeedback … |
| OLE | 58 | 20 | 16: Updated, Click, DblClick, DragDrop, DragOver, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove … |
| UserControl | 62 | 22 | 35: Initialize, DragDrop, DragOver, Resize, Click, DblClick, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown … |
| PropertyPage | 48 | 16 | 24: SelectionChanged, DragDrop, DragOver, Click, DblClick, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown, MouseMove … |
| UserDocument | 56 | 21 | 32: Initialize, DragDrop, DragOver, Resize, Click, DblClick, GotFocus, KeyDown, KeyPress, KeyUp, LostFocus, MouseDown … |
| Global | 7 | 7 | 0 |
| VBControlExtender | 0 | 34 | 10: GotFocus, LostFocus, DeprecatedEvent1, DeprecatedEvent2, DeprecatedEvent3, DeprecatedEvent4, DragDrop, DragOver, Validate, ObjectEvent |

## Common Controls (MSCOMCTL.OCX)

| Control | Properties | Methods | Events |
|---|---|---|---|
| TabStrip | 25 | 3 | 14 |
| Toolbar | 21 | 5 | 15 |
| StatusBar | 10 | 2 | 13 |
| ProgressBar | 12 | 2 | 10 |
| TreeView | 25 | 5 | 20 |
| ListView | 38 | 6 | 19 |
| ImageList | 7 | 1 | 0 |
| Slider | 20 | 3 | 15 |
| ImageCombo | 18 | 3 | 12 |

## Common Controls 2 (MSCOMCT2.OCX)

| Control | Properties | Methods | Events |
|---|---|---|---|
| Animation | 7 | 5 | 11 |
| UpDown | 14 | 1 | 12 |
| MonthView | 34 | 4 | 18 |
| DTPicker | 25 | 2 | 20 |
| FlatScrollBar | 12 | 1 | 2 |

## CoolBar (COMCT332.OCX)

| Control | Properties | Methods | Events |
|---|---|---|---|
| CoolBar | 18 | 2 | 13 |

## CommonDialog (COMDLG32.OCX)

| Control | Properties | Methods | Events |
|---|---|---|---|
| CommonDialog | 30 | 6 | 0 |

## MSFlexGrid (MSFLXGRD.OCX)

| Control | Properties | Methods | Events |
|---|---|---|---|
| MSFlexGrid | 89 | 5 | 20 |

## RichTextBox (RICHTX32.OCX)

| Control | Properties | Methods | Events |
|---|---|---|---|
| RichTextBox | 43 | 9 | 16 |

## SSTab (TABCTL32.OCX)

| Control | Properties | Methods | Events |
|---|---|---|---|
| SSTab | 24 | 1 | 14 |

## VBA runtime library (MSVBVM60.DLL)

Every function of the VBA modules maps to an MSVBVM60 export (`rtc*`), kept in `VBA.json` as `export`. Names `_B_str_X` / `_B_var_X` are the `X$` and `X` forms of the same function.

| Module | Functions | Examples |
|---|---|---|
| Constants | 0 |  |
| Strings | 55 | Asc → rtcAnsiValueBstr, _B_str_Chr → rtcBstrFromAnsi, _B_var_Chr → rtcVarBstrFromAnsi |
| Conversion | 25 | _B_str_Hex → rtcHexBstrFromVar, _B_var_Hex → rtcHexVarFromVar, _B_str_Oct → rtcOctBstrFromVar |
| FileSystem | 20 | ChDir → rtcChangeDir, ChDrive → rtcChangeDrive, EOF → rtcEndOfFile |
| DateTime | 26 | _B_var_Date → rtcGetDateVar, _B_str_Date → rtcSetDateBstr, _B_var_Date → rtcSetDateVar |
| Information | 15 | Erl → rtcGetErl, Err → rtcErrObj, IMEStatus → rtcIMEStatus |
| Interaction | 23 | AppActivate → rtcAppActivate, Beep → rtcBeep, CreateObject → rtcCreateObject2 |
| Math | 12 | Abs → rtcAbsVar, Atn → rtcAtn, Cos → rtcCos |
| Financial | 13 | SLN → rtcSLN, SYD → rtcSYD, DDB → rtcDDB |
| _HiddenModule | 9 | Array → rtcArray, _B_str_InputB → rtcInputCount, _B_var_InputB → rtcInputCountVar |

Enumerations: VbVarType, VbMsgBoxStyle, VbMsgBoxResult, VbFileAttribute, VbStrConv, VbDayOfWeek, VbFirstWeekOfYear, VbIMEStatus, VbAppWinStyle, VbCompareMethod, VbCalendar, VbDateTimeFormat, VbTriState, VbCallType. Objects: ErrObject, Collection.

## VBRUN constants

62 enumerations (593 constants), e.g. PaletteModeConstants, AsyncTypeConstants, AsyncReadConstants, AsyncStatusCodeConstants, LogEventTypeConstants, QueryUnloadConstants, ZOrderConstants, ShiftConstants, MouseButtonConstants, ApplicationStartConstants …

## IDE (VB6IDE.DLL)

| Menu | Commands |
|---|---|
| File | New Project, Open Project..., Add Project..., Remove Project, Save Project Group, Save Project Group As..., Save <File Name>, Save <File Name> As..., Print..., Print Setup..., Make <Project>..., Make Project Group..., Exit, Save Selection, Save Change Script |
| Edit | Undo, Redo, Cut, Copy, Paste, Paste Link, Delete, Select All, Find..., Find Next, Replace..., Indent, Outdent, Insert File..., List Properties/Methods, List Constants, Quick Info, Parameter Info, Complete Word, Bookmarks |
| View | Code, Object, Definition, Last Position, Object Browser, Immediate Window, Locals Window, Watch Window, Call Stack..., Project Explorer, Properties Window, Property Pages, Toolbox, Tab Order, Color Palette, Toolbars, Form Layout Window |
| Project | Add Procedure..., Module, Class Module, Add Form, Add MDI Form, Add Module, Add Class Module, Add User Control, Add Property Page, Add User Document, More ActiveX Designers..., Add File..., Remove <Item>, References..., Components..., <Project Name> Properties..., User Control, Property Page, User Document, Add  |
| Format | Align, Make Same Size, Size to Grid, Horizontal Spacing, Vertical Spacing, Center in Form, Arrange Buttons, Order, Lock Controls |
| Debug | Step Into, Step Over, Step Out, Run To Cursor, Add Watch..., Edit Watch..., Quick Watch..., Disable All Watches, Toggle Breakpoint, Clear All Breakpoints, Disable All Breakpoints, Set Next Statement, Show Next Statement |
| Run | Start With Full Compile, Break, End, Restart, Start |
| Tools | Procedure Attributes..., Add-In Manager..., Menu Editor..., References..., Get, Check Out, Check In, Undo Check Out, Options..., Visual Data Manager..., Report Designer... |
| Window | Split, Tile Horizontally, Tile Vertically, Cascade, Arrange Icons, More Windows... |
| Help | Contents..., Index..., Technical Support, About Microsoft Visual Basic..., Microsoft on the Web, Search... |

Toolbars: Standard, Edit, Debug, Form Editor. Context menus: Project Window Insert, Toggle, Order, ActiveX Designer, Toolbars, Code Window, Code Window (Break), Watch Window, Immediate Window, Locals Window, Forms, Controls, Project Window, Project Window (Break), Object Browser, Toolbox, Toolbox Group, Property Browser, Color Palette, Toolbar, Project Window Project, Project Window Form Folder, Project Window Module/Class Folder, Project Window Related Documents Folder, Standard, Edit, Debug, UserForm, Form Editor, Project Window.

Dialogs: About Microsoft Visual Basic, Add Procedure, Add Watch, Add-In Manager, Advanced Optimizations, Call Stack, Choose Class, Command Line Options, Component Busy, Connect Property Pages, Context Help, Custom Control Upgrade, Custom Controls, Data Bindings, Define Color, Edit Watch, File New, Find, Member Options, Menu Editor, Open, Print, Procedure Attributes, Project Password, Quick Watch, References, Replace, View Procedure, Windows.

Keyboard shortcuts (main table): Ctrl+S, Ctrl+O, Ctrl+N, Ctrl+E, Shift+F4, Alt+F5, Alt+F8, Ctrl+A, F2, Shift+F2, Ctrl+Shift+F2, Shift+F5, F5, Ctrl+Break, Ctrl+F5, F7, Shift+F7, F8, Ctrl+Shift+F8, Shift+F8, Ctrl+F8, F9, Shift+F9, Ctrl+F9, Ctrl+Shift+F9, Ctrl+D, Ctrl+F, F3, Shift+F3, Ctrl+F3, Ctrl+Shift+F3, Ctrl+J, Ctrl+Shift+J, Ctrl+I, Ctrl+Shift+I, Ctrl+Space, Ctrl+G, Ctrl+H, Ctrl+J, Ctrl+K, Ctrl+P, Ctrl+R, Ctrl+W, Alt+Q, Ctrl+L, Ctrl+Tab, Ctrl+Shift+Tab, Ctrl+T.

## Run-time errors

89 VBA errors (e.g. 13 Type mismatch, 91 Object variable or With block variable not set) and 139 VB errors of forms and controls (e.g. 340 Control array element '|1' doesn't exist, 360 Object already loaded, 364 Object was unloaded), with the messages MSVBVM60 gives `Err.Description` (`runtime_errors.json`). MSVBVM60 keeps the VBA messages as string ids 10000+i in ascending error order and the VB messages as string id 1000 + error number.

## File formats

`tools/rea/check_samples.py` on the VB98 samples of the MSDN library: every .frm/.ctl/.dob read and written back byte for byte; `.frx` pictures (BMP, ICO, WMF), ImageList images (`StdPicture` CLSID prefix) and `$` strings (u32 length) and short `Text` (u8 length) decoded on 75 files. `List`/`ItemData` taken from a form saved by VB6 itself (a test form saved in the VB6 IDE): `[u16 count][u16 max length]` then `[u16 length][text]` per item; ItemData numbers are stored as text. Bindings, OleObjectBlob and PropertyPages records are kept as raw bytes.

## Open items

* Types from libraries not loaded (msdatsrc, MSSTDFMT, DAO) appear as `file:{GUID}` in the JSON.
