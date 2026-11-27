VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect
   ClientHeight    =   9945
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   "Logs the events and the objects of the IDE"
   DisplayName     =   "EvLog"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VBI As VBIDE.VBE
Public WithEvents PE As VBProjectsEvents
Public WithEvents CE As VBComponentsEvents
Public WithEvents RE As ReferencesEvents
Public WithEvents FE As FileControlEvents
Public WithEvents XE As VBControlsEvents
Public WithEvents SE As SelectedVBControlsEvents
Public WithEvents MH As CommandBarEvents
Dim mCtl As Office.CommandBarControl

Function Sep() As String
    If InStr(App.Path, "/") > 0 Then Sep = "/" Else Sep = "\"
End Function

Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & Sep & "evlog.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Function Flat(ByVal s As String) As String
    Flat = Replace(Replace(s, vbCrLf, "|"), vbLf, "|")
End Function

Function FileOf(ByVal s As String) As String
    Dim i As Long
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) = "\" Or Mid$(s, i, 1) = "/" Then Exit For
    Next
    FileOf = Mid$(s, i + 1)
End Function

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo bad
    Set VBI = Application
    Log "connect " & ConnectMode & " " & TypeName(Application) & " " & TypeName(AddInInst) & " " & AddInInst.ProgId & " " & AddInInst.Connect
    Log "vbe " & VBI.Version & " " & VBI.Name & " projects " & VBI.VBProjects.Count
    Set PE = VBI.Events.VBProjectsEvents
    Set CE = VBI.Events.VBComponentsEvents(Nothing)
    Set RE = VBI.Events.ReferencesEvents(Nothing)
    Set FE = VBI.Events.FileControlEvents(Nothing)
    Set mCtl = VBI.CommandBars("Add-Ins").Controls.Add(1)
    mCtl.Caption = "EvLog Run"
    Set MH = VBI.Events.CommandBarEvents(mCtl)
    Log "bar " & TypeName(mCtl) & " " & mCtl.Caption & " " & mCtl.Type & " " & TypeName(mCtl.Parent) & " " & mCtl.Parent.Name
    Exit Sub
bad:
    Log "connect error " & Err.Number & " " & Err.Description
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    Log "startupcomplete"
    RunTests
End Sub

Private Sub MH_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Log "click " & CommandBarControl.Caption & " " & handled & " " & CancelDefault
    RunTests
End Sub

Sub RunTests()
    Dim p As VBProject, c As VBComponent, r As Reference, i As Long, k As Long, s As String
    Dim f As VBComponent, ctl As VBControl, cm As CodeModule
    On Error GoTo bad
    Set p = VBI.ActiveVBProject
    Log "proj " & p.Name & " type " & p.Type & " comps " & p.VBComponents.Count & " refs " & p.References.Count & " file " & FileOf(p.FileName)
    For Each r In p.References
        Log "ref " & r.Name & " " & r.Guid & " " & r.Major & "." & r.Minor & " type " & r.Type & " builtin " & r.BuiltIn & " broken " & r.IsBroken
    Next
    For Each c In p.VBComponents
        Log "comp " & c.Name & " " & c.Type & " lines " & c.CodeModule.CountOfLines
    Next
    s = ""
    For i = 1 To 5
        s = s & VBI.CommandBars("Menu Bar").Controls(i).Caption & ";"
    Next
    Log "menubar " & s & " tools " & TypeName(VBI.CommandBars("Tools"))

    Log "-- add module"
    Set c = p.VBComponents.Add(vbext_ct_StdModule)
    Log "added " & c.Name & " " & c.Type & " lines " & c.CodeModule.CountOfLines & " " & Flat(c.CodeModule.Lines(1, c.CodeModule.CountOfLines))
    Set cm = c.CodeModule
    cm.AddFromString "Dim x As Integer" & vbCrLf & vbCrLf & "' about f" & vbCrLf & "Sub f()" & vbCrLf & "    x = 1" & vbCrLf & "End Sub" & vbCrLf & "Function g() As Long" & vbCrLf & "End Function"
    Log "code " & cm.CountOfLines & " decl " & cm.CountOfDeclarationLines & " | " & Flat(cm.Lines(1, cm.CountOfLines))
    Log "proc f start " & cm.ProcStartLine("f", vbext_pk_Proc) & " body " & cm.ProcBodyLine("f", vbext_pk_Proc) & " count " & cm.ProcCountLines("f", vbext_pk_Proc)
    Log "proc g start " & cm.ProcStartLine("g", vbext_pk_Proc) & " body " & cm.ProcBodyLine("g", vbext_pk_Proc) & " count " & cm.ProcCountLines("g", vbext_pk_Proc)
    s = cm.ProcOfLine(6, k)
    Log "procofline 6 " & s & " " & k
    Dim sl As Long, sc As Long, el As Long, ec As Long
    sl = 1: sc = 1: el = -1: ec = -1
    Log "find " & cm.Find("x = 1", sl, sc, el, ec) & " " & sl & " " & sc & " " & el & " " & ec
    cm.InsertLines 2, "' inserted"
    cm.ReplaceLine 3, "Dim y As Long"
    cm.DeleteLines 1, 1
    Log "edited " & cm.CountOfLines & " | " & Flat(cm.Lines(1, cm.CountOfLines))
    Log "-- rename"
    c.Name = "ModEv"
    Log "renamed " & c.Name

    Log "-- add form"
    Set f = p.VBComponents.Add(vbext_ct_VBForm)
    Log "form " & f.Name & " " & f.Type & " caption " & f.Properties("Caption").Value & " props " & (f.Properties.Count > 10)
    f.Properties("Caption").Value = "EvLog form"
    Log "caption " & f.Properties("Caption").Value
    Set XE = VBI.Events.VBControlsEvents(p, f.Designer)
    Set SE = VBI.Events.SelectedVBControlsEvents(p, f.Designer)
    Log "designer " & TypeName(f.Designer) & " controls " & f.Designer.VBControls.Count
    Dim o As Object
    Set o = f.Designer.VBControls.Add("VB.CommandButton")
    If o Is Nothing Then Set o = f.Designer.VBControls.Add("VB.CommandButton", Nothing, False)
    Log "added control " & TypeName(o)
    Set ctl = o
    Log "control " & TypeName(ctl) & " " & ctl.Properties("Name").Value & " " & ctl.ProgId & " caption " & ctl.Properties("Caption").Value
    ctl.Properties("Caption").Value = "Go"
    ctl.Properties("Name").Value = "cmdGo"
    ctl.Properties("Left").Value = 240
    Log "control " & ctl.Properties("Name").Value & " " & ctl.Properties("Caption").Value & " " & ctl.Properties("Left").Value & " count " & f.Designer.VBControls.Count
    i = f.CodeModule.CreateEventProc("Click", "cmdGo")
    Log "eventproc " & i & " | " & Flat(f.CodeModule.Lines(1, f.CodeModule.CountOfLines))
    f.Designer.VBControls.Remove ctl
    Log "control removed " & f.Designer.VBControls.Count
    Set XE = Nothing
    Set SE = Nothing

    Log "-- references"
    Set r = p.References.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
    Log "addref " & r.Name & " " & p.References.Count
    p.References.Remove r
    Log "removed ref " & p.References.Count

    Log "-- remove"
    p.VBComponents.Remove f
    p.VBComponents.Remove c
    Log "comps " & p.VBComponents.Count
    Log "-- import"
    Set c = p.VBComponents.AddFile(App.Path & Sep & "host" & Sep & "ModEv.bas")
    Log "imported " & c.Name & " " & c.Type & " " & FileOf(c.FileNames(1)) & " comps " & p.VBComponents.Count
    Log "code " & c.CodeModule.CountOfLines & " decl " & c.CodeModule.CountOfDeclarationLines
    Log "-- projects"
    Dim p2 As VBProject
    Set p2 = VBI.VBProjects.Add(vbext_pt_StandardExe)
    Log "p2 " & p2.Name & " " & p2.VBComponents.Count & " projects " & VBI.VBProjects.Count & " active " & VBI.ActiveVBProject.Name
    p2.Name = "Other"
    Set VBI.ActiveVBProject = p
    Log "active " & VBI.ActiveVBProject.Name
    Log "done"
    VBI.VBProjects.Remove p2
    Log "projects " & VBI.VBProjects.Count
    Log "done2"
    Exit Sub
bad:
    Log "error " & Err.Number & " " & Err.Description & " (" & Erl & ")"
    Resume Next
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    Log "beginshutdown"
End Sub

Private Sub AddinInstance_OnAddInsUpdate(custom() As Variant)
    Log "addinsupdate"
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    mCtl.Delete
    Log "disconnect " & RemoveMode
End Sub

Private Sub PE_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Log "PE added " & VBProject.Name
End Sub
Private Sub PE_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    Log "PE removed " & VBProject.Name
End Sub
Private Sub PE_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    Log "PE renamed " & VBProject.Name & " " & OldName
End Sub
Private Sub PE_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    Log "PE activated " & VBProject.Name
End Sub
Private Sub CE_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    Log "CE added " & VBComponent.Name & " " & VBComponent.Type
End Sub
Private Sub CE_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    Log "CE removed " & VBComponent.Name
End Sub
Private Sub CE_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    Log "CE renamed " & VBComponent.Name & " " & OldName
End Sub
Private Sub CE_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    Log "CE selected " & VBComponent.Name
End Sub
Private Sub CE_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    Log "CE activated " & VBComponent.Name
End Sub
Private Sub CE_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
    Log "CE reloaded " & VBComponent.Name
End Sub
Private Sub RE_ItemAdded(ByVal Reference As VBIDE.Reference)
    Log "RE added " & Reference.Name
End Sub
Private Sub RE_ItemRemoved(ByVal Reference As VBIDE.Reference)
    Log "RE removed " & Reference.Name
End Sub
Private Sub FE_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    Log "FE afteradd " & FileType & " " & FileOf(FileName)
End Sub
Private Sub FE_AfterRemoveFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    Log "FE afterremove " & FileType & " " & FileOf(FileName)
End Sub
Private Sub FE_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
    Log "FE afterwrite " & FileType & " " & FileOf(FileName) & " " & Result
End Sub
Private Sub FE_AfterChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String)
    Log "FE afterchange " & FileType & " " & FileOf(NewName) & " " & FileOf(OldName)
End Sub
Private Sub FE_AfterCloseFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal WasDirty As Boolean)
    Log "FE afterclose " & FileType & " " & FileOf(FileName) & " " & WasDirty
End Sub
Private Sub FE_RequestWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileName As String, Cancel As Boolean)
    Log "FE requestwrite " & FileOf(FileName) & " " & Cancel
End Sub
Private Sub FE_RequestChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String, Cancel As Boolean)
    Log "FE requestchange " & FileType & " " & FileOf(NewName) & " " & FileOf(OldName)
End Sub
Private Sub FE_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
    Log "FE dogetnew " & FileType & " " & FileOf(NewName) & " " & FileOf(OldName)
End Sub
Private Sub XE_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    Log "XE added " & VBControl.Properties("Name").Value
End Sub
Private Sub XE_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    Log "XE removed " & VBControl.Properties("Name").Value
End Sub
Private Sub XE_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
    Log "XE renamed " & VBControl.Properties("Name").Value & " " & OldName & " " & OldIndex
End Sub
Private Sub SE_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    Log "SE added " & VBControl.Properties("Name").Value
End Sub
Private Sub SE_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    Log "SE removed " & VBControl.Properties("Name").Value
End Sub
