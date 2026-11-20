"""The ActiveX control libraries (.ocx) that the forms runtime has (runtime/forms/ocx*.inc): what the
Components dialog lists, what the Toolbox adds for each, and how the designer writes a new one.

Each library is named in the .vbp and at the top of a .frm by its type library:
    Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX          (.vbp)
    Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"     (.frm)
"""


class OcxControl:
    """one control of a library: its class, the base of new names, the default size (twips), the
    _Version VB6 writes, and the properties a new one gets"""

    def __init__(self, cls, base, w, h, version=393216, invisible=False, container=False, extra=None, bag=None,
                 streamid=None, extent="_"):
        self.cls, self.base, self.w, self.h = cls, base, w, h
        self.extent = extent                     # "_ExtentX" ...; the WebBrowser writes "ExtentX" (no _Version)
        self.streamid = streamid                 # _StreamID before _Version (the DataRepeater)
        self.version, self.invisible, self.container = version, invisible, container
        self.extra = extra or []
        # the properties the control writes itself, after _Version, in its own order and without the
        # names of enum values (VB6 oracle d_dev); None: in the alphabetical order of the extender
        self.bag = bag


class OcxLib:
    def __init__(self, key, guid, version, file, title, controls):
        self.key, self.guid, self.version, self.file, self.title = key, guid, version, file, title
        self.controls = controls

    @property
    def vbp_line(self):
        return f"{self.guid}#{self.version}#0; {self.file}"

    @property
    def frm_line(self):
        return f'Object = "{self.guid}#{self.version}#0"; "{self.file}"'


LIBS = [
    OcxLib("MSComDlg", "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}", "1.2", "COMDLG32.OCX",
           "Microsoft Common Dialog Control 6.0 (SP6)",
           [OcxControl("CommonDialog", "CommonDialog", 480, 480, invisible=True)]),
    OcxLib("MSComctlLib", "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}", "2.2", "MSCOMCTL.OCX",
           "Microsoft Windows Common Controls 6.0 (SP6)",
           [OcxControl("TabStrip", "TabStrip", 2895, 1815),
            OcxControl("Toolbar", "Toolbar", 4800, 420, extra=[("Align", 1), ("Appearance", 1)]),
            OcxControl("StatusBar", "StatusBar", 4800, 375, extra=[("Align", 2)]),
            OcxControl("ProgressBar", "ProgressBar", 2415, 255, extra=[("Appearance", 1)]),
            OcxControl("TreeView", "TreeView", 2295, 1695, version=393217, extra=[("Style", 7), ("Appearance", 1)]),
            OcxControl("ListView", "ListView", 2535, 1695, version=393217,
                       extra=[("LabelWrap", True), ("HideSelection", True), ("ForeColor", -2147483640),
                              ("BackColor", -2147483643), ("BorderStyle", 1), ("Appearance", 1), ("NumItems", 0)]),
            OcxControl("ImageList", "ImageList", 480, 480, invisible=True,
                       extra=[("BackColor", -2147483643), ("MaskColor", 12632256)]),
            OcxControl("Slider", "Slider", 2415, 615),
            OcxControl("ImageCombo", "ImageCombo", 2055, 330, extra=[("ForeColor", -2147483640),
                                                                     ("BackColor", -2147483643)])]),
    OcxLib("MSComCtl2", "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}", "2.0", "MSCOMCT2.OCX",
           "Microsoft Windows Common Controls-2 6.0 (SP6)",
           [OcxControl("Animation", "Animation", 1215, 735, extra=[("FullWidth", 81), ("FullHeight", 49)]),
            OcxControl("UpDown", "UpDown", 255, 495),
            OcxControl("MonthView", "MonthView", 2700, 2370, extra=[("Appearance", 1)]),
            OcxControl("DTPicker", "DTPicker", 1695, 315, extra=[("Format", 1)]),
            OcxControl("FlatScrollBar", "FlatScrollBar", 2415, 255, extra=[("Arrows", 65536),
                                                                            ("Orientation", 1376257)])]),
    OcxLib("TabDlg", "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}", "1.1", "TABCTL32.OCX",
           "Microsoft Tabbed Dialog Control 6.0 (SP6)",
           [OcxControl("SSTab", "SSTab", 3615, 2175, container=True,
                       extra=[("Tabs", 3), ("TabsPerRow", 3), ("TabHeight", 520),
                              ("TabCaption(0)", "Tab 0"), ("TabCaption(1)", "Tab 1"), ("TabCaption(2)", "Tab 2")])]),
    OcxLib("MSFlexGridLib", "{5E9E78A0-531B-11CF-91F6-C2863C385E30}", "1.0", "MSFLXGRD.OCX",
           "Microsoft FlexGrid Control 6.0 (SP6)",
           [OcxControl("MSFlexGrid", "MSFlexGrid", 2775, 1455)]),
    # the grid of the Data Form Wizard's "MS HFlexGrid" layout (phase 9.2b): the runtime's flex grid, bound
    OcxLib("MSHierarchicalFlexGridLib", "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}", "6.0", "MSHFLXGD.OCX",
           "Microsoft Hierarchical FlexGrid Control 6.0 (SP4) (OLEDB)",
           [OcxControl("MSHFlexGrid", "MSHFlexGrid", 2775, 1455)]),
    OcxLib("RichTextLib", "{3B7C8863-D78F-101B-B9B5-04021C009402}", "1.2", "RICHTX32.OCX",
           "Microsoft Rich Textbox Control 6.0 (SP6)",
           [OcxControl("RichTextBox", "RichTextBox", 2055, 1215, version=393217, extra=[("Enabled", True)])]),
    OcxLib("MSChart20Lib", "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}", "2.0", "MSCHRT20.OCX",
           "Microsoft Chart Control 6.0 (SP4) (OLEDB)",
           [OcxControl("MSChart", "MSChart", 4095, 3015)]),
    OcxLib("MSAdodcLib", "{67397AA1-7FB1-11D0-B148-00A0C922E820}", "6.0", "MSADODC.OCX",
           "Microsoft ADO Data Control 6.0 (SP6) (OLEDB)",
           [OcxControl("Adodc", "Adodc", 2415, 330,
                       extra=[("ConnectMode", 0), ("CursorLocation", 3), ("IsolationLevel", -1),
                              ("ConnectionTimeout", 15), ("CommandTimeout", 30), ("CursorType", 3), ("LockType", 3),
                              ("CommandType", 8), ("CursorOptions", 0), ("CacheSize", 50), ("MaxRecords", 0),
                              ("BOFAction", 0), ("EOFAction", 0), ("ConnectStringType", 1), ("Appearance", 1),
                              ("BackColor", -2147483643), ("ForeColor", -2147483640), ("Orientation", 0),
                              ("Enabled", -1), ("Connect", ""), ("OLEDBString", ""), ("OLEDBFile", ""),
                              ("DataSourceName", ""), ("OtherAttributes", ""), ("UserName", ""), ("Password", ""),
                              ("RecordSource", ""), ("Caption", "$NAME")])]),
    OcxLib("MSDataGridLib", "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}", "1.0", "MSDATGRD.OCX",
           "Microsoft DataGrid Control 6.0 (SP6) (OLEDB)",
           [OcxControl("DataGrid", "DataGrid", 6495, 3015, extra=[("HeadLines", 1), ("RowHeight", 15)])]),
    OcxLib("MSDataListLib", "{F0D2F211-CCB0-11D0-A316-00AA00688B10}", "1.0", "MSDATLST.OCX",
           "Microsoft DataList Controls 6.0 (SP6) (OLEDB)",
           [OcxControl("DataList", "DataList", 1935, 1425),
            OcxControl("DataCombo", "DataCombo", 1935, 315, extra=[("Text", "$NAME")])]),
    OcxLib("MCI", "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}", "1.1", "mci32.Ocx",
           "Microsoft Multimedia Control 6.0",
           [OcxControl("MMControl", "MMControl", 3540, 495, extra=[("DeviceType", ""), ("FileName", "")],
                       bag=["Frames", "MousePointer", "OLEDropMode", "Orientation", "BorderStyle", "RecordMode",
                            "UpdateInterval",
                            "PrevEnabled", "NextEnabled", "PlayEnabled", "PauseEnabled", "BackEnabled",
                            "StepEnabled", "StopEnabled", "RecordEnabled", "EjectEnabled", "AutoEnable", "Silent",
                            "PrevVisible", "NextVisible", "PlayVisible", "PauseVisible", "BackVisible",
                            "StepVisible", "StopVisible", "RecordVisible", "EjectVisible",
                            "Shareable", "Enabled", "DeviceType", "FileName"])]),
    OcxLib("MSCommLib", "{648A5603-2C6E-101B-82B6-000000000014}", "1.1", "MSComm32.Ocx",
           "Microsoft Comm Control 6.0",
           [OcxControl("MSComm", "MSComm", 570, 570, invisible=True, extra=[("DTREnable", True)],
                       bag=["CommPort", "DTREnable", "Handshaking", "InBufferSize", "InputLen", "NullDiscard",
                            "OutBufferSize", "ParityReplace", "RThreshold", "RTSEnable", "BaudRate",
                            "ParitySetting", "DataBits", "StopBits", "SThreshold", "EOFEnable", "InputMode"])]),
    OcxLib("MSWinsockLib", "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}", "1.0", "MSWINSCK.OCX",
           "Microsoft Winsock Control 6.0 (SP6)",
           [OcxControl("Winsock", "Winsock", 420, 420, invisible=True,
                       bag=["Protocol", "RemoteHost", "RemotePort", "LocalPort"])]),
    OcxLib("InetCtlsObjects", "{48E59290-9880-11CF-9754-00AA00C00908}", "1.0", "MSINET.OCX",
           "Microsoft Internet Transfer Control 6.0 (SP6)",
           [OcxControl("Inet", "Inet", 570, 570, invisible=True,
                       bag=["AccessType", "Protocol", "RemoteHost", "RemotePort", "URL", "Document", "UserName",
                            "Password", "Proxy", "RequestTimeout"])]),
    OcxLib("MSMAPI", "{20C62CAE-15DA-101B-B9A8-444553540000}", "1.1", "msmapi32.Ocx",
           "Microsoft MAPI Controls 6.0",
           # as the VB6 IDE writes them (VB6 oracle f_mapi5): the numbers and Booleans first, then the strings
           [OcxControl("MAPISession", "MAPISession", 570, 570, invisible=True,
                       bag=["DownloadMail", "LogonUI", "NewSession", "Password", "UserName"]),
            OcxControl("MAPIMessages", "MAPIMessages", 570, 570, invisible=True,
                       bag=["AddressEditFieldCount", "AddressModifiable", "AddressResolveUI", "FetchSorted",
                            "FetchUnreadOnly", "AddressCaption", "AddressLabel", "FetchMsgType"])]),
    OcxLib("MSMask", "{C932BA88-4374-101B-A56C-00AA003668DC}", "1.1", "MSMASK32.OCX",
           "Microsoft Masked Edit Control 6.0 (SP6)",
           [OcxControl("MaskEdBox", "MaskEdBox", 1095, 375, extra=[("PromptChar", "_")],
                       bag=["ClipMode", "Appearance", "BorderStyle", "BackColor", "ForeColor", "HideSelection",
                            "PromptInclude", "AllowPrompt", "AutoTab", "Enabled", "MousePointer", "OLEDragMode",
                            "OLEDropMode", "MaxLength", "Font", "Format", "Mask", "PromptChar"])]),
    OcxLib("MSDBGrid", "{00028C01-0000-0000-0000-000000000046}", "1.0", "DBGRID32.OCX",
           "Microsoft Data Bound Grid Control 5.0 (SP3)",
           # its design-time properties are in the OleObjectBlob of the .frx (lbvc/project/oleblob.py)
           [OcxControl("DBGrid", "DBGrid", 6495, 3015, version=None)]),
    OcxLib("MSDBCtls", "{FAEEE763-117E-101B-8933-08002B2F4F5A}", "1.1", "DBLIST32.OCX",
           "Microsoft Data Bound List Controls 6.0",
           [OcxControl("DBList", "DBList", 1935, 1035, bag=["ListField", "BoundColumn"]),
            OcxControl("DBCombo", "DBCombo", 1935, 315, extra=[("Text", "$NAME")],
                       bag=["ListField", "BoundColumn", "Style", "Text"])]),
    OcxLib("ComCtl3", "{38911DA0-E448-11D0-84A3-00DD01104159}", "1.1", "COMCT332.OCX",
           "Microsoft Windows Common Controls-3 6.0 (SP5)",
           # its bands: flat properties of the band number (vbprops.cool_bands); _Version is a string
           [OcxControl("CoolBar", "CoolBar", 4800, 420, version="6.7.9839", container=True,
                       extra=[("Align", 1), ("BandCount", 3), ("_CBWidth", "$W"), ("_CBHeight", "$H"),
                              ("MinHeight1", 360), ("NewRow1", False), ("MinHeight2", 360), ("NewRow2", False),
                              ("MinHeight3", 360), ("NewRow3", False)])]),
    OcxLib("MSRDC", "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}", "2.0", "MSRDC20.OCX",
           "Microsoft RemoteData Control 6.0 (SP6)",
           # as the VB6 IDE writes it (VB6 oracle f_rdo3): ResultsetType as RecordsetType, SQL as RecordSource,
           # Appearance twice
           [OcxControl("MSRDC", "MSRDC", 2415, 330,
                       extra=[("Options", 0), ("CursorDriver", 0), ("BOFAction", 0), ("EOFAction", 0),
                              ("RecordsetType", 1), ("LockType", 3), ("QueryType", 0), ("Prompt", 0),
                              ("Appearance", 1), ("QueryTimeout", 30), ("RowsetSize", 100), ("LoginTimeout", 15),
                              ("KeysetSize", 0), ("MaxRows", 0), ("ErrorThreshold", -1), ("BatchSize", 15),
                              ("BackColor", -2147483643), ("ForeColor", -2147483640), ("Enabled", -1),
                              ("ReadOnly", 0), ("Appearance", -1), ("DataSourceName", ""), ("RecordSource", ""),
                              ("UserName", ""), ("Password", ""), ("Connect", ""), ("LogMessages", ""),
                              ("Caption", "$NAME")])]),
    OcxLib("MSDataRepeaterLib", "{5C8CED40-8909-11D0-9483-00A0C91110ED}", "1.0", "MSDatRep.Ocx",
           "Microsoft DataRepeater Control 6.0 (OLEDB)",
           # as the VB6 IDE writes it (VB6 oracle f_drep3): _StreamID before _Version, the properties that are not
           # the defaults in this order, then RepeatedControlName and the RepeaterBinding(i) blocks (vbprops.set_drep)
           [OcxControl("DataRepeater", "DataRepeater", 3615, 2895, streamid=-1412567295, extra=[("Caption", "$NAME")],
                       bag=["RowIndicator", "Enabled", "Appearance", "BackColor", "ForeColor", "BorderStyle",
                            "MousePointer", "RightToLeft", "ScrollBars", "RowDividerStyle", "CaptionStyle", "Caption",
                            "Font", "IntegralHeight"])]),
    OcxLib("SysInfoLib", "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}", "1.0", "SysInfo.Ocx",
           "Microsoft SysInfo Control 6.0",
           [OcxControl("SysInfo", "SysInfo", 570, 570, invisible=True)]),
    # phase 9.2: the WebBrowser of the Application Wizard (runtime/forms/ocx_web.inc, on QTextBrowser); as the VB6 IDE
    # writes it (VB6 oracle DEV-9/aw3 frmBrowser): ExtentX / ExtentY without "_", no _Version, these properties
    OcxLib("SHDocVwCtl", "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}", "1.1", "ieframe.dll",
           "Microsoft Internet Controls",
           [OcxControl("WebBrowser", "WebBrowser", 3735, 2535, version=None, extent="",
                       extra=[("ViewMode", 0), ("Offline", 0), ("Silent", 0), ("RegisterAsBrowser", 0),
                              ("RegisterAsDropTarget", 1), ("AutoArrange", False), ("NoClientEdge", False),
                              ("AlignLeft", False), ("NoWebView", False), ("HideFileNames", False),
                              ("SingleClick", False), ("SingleSelection", False), ("NoFolders", False),
                              ("Transparent", False), ("ViewID", "{0057D0E0-3573-11CF-AE69-08002B2E1262}"),
                              ("Location", "")],
                       bag=["ExtentX", "ExtentY", "ViewMode", "Offline", "Silent", "RegisterAsBrowser",
                            "RegisterAsDropTarget", "AutoArrange", "NoClientEdge", "AlignLeft", "NoWebView",
                            "HideFileNames", "SingleClick", "SingleSelection", "NoFolders", "Transparent", "ViewID",
                            "Location"])]),
    OcxLib("PicClip", "{27395F88-0C0C-101B-A3C9-08002B2F49FB}", "1.1", "PICCLP32.OCX",
           "Microsoft PictureClip Control 6.0 (SP6)",
           [OcxControl("PictureClip", "PictureClip", 1215, 1215, extra=[("_StockProps", 0)])]),
]
LIB_BY_KEY = {l.key.lower(): l for l in LIBS}
LIB_BY_GUID = {l.guid.upper(): l for l in LIBS}
# the class of a control (Lib.Cls) -> (library, control)
CONTROL = {f"{l.key}.{c.cls}": (l, c) for l in LIBS for c in l.controls}
CONTROL_LOW = {k.lower(): v for k, v in CONTROL.items()}


def lib_of_object_line(value):
    """'{GUID}#2.2#0; MSCOMCTL.OCX' (.vbp) or the .frm line -> the library, or None"""
    import re
    m = re.search(r"\{[0-9A-Fa-f-]{36}\}", value or "")
    return LIB_BY_GUID.get(m.group(0).upper()) if m else None


def control_of(cls):
    """'MSComctlLib.TreeView' -> (library, control) or None"""
    return CONTROL_LOW.get((cls or "").lower())


def extent(twips):
    """_ExtentX / _ExtentY: HIMETRIC (0.01 mm)"""
    return int(round(twips * 2540 / 1440))


def is_invisible(cls):
    lc = control_of(cls)
    return bool(lc and lc[1].invisible)


def is_container(cls):
    lc = control_of(cls)
    return bool(lc and lc[1].container)
