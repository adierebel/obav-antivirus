Attribute VB_Name = "bsFORM"
Option Explicit
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Type OSVERSIONINFO
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion As String * 128
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

Dim tmpCMD As String
Public hMMTimer As Long
Dim cImgList  As gComCtl
Dim enumproC() As ENUMERATE_PROCESSES_OUTPUT
Private Sub prepScan()
scAn = True
jmlVIR = 0
scannedF = 0
Perc = 0
registriBad = 0
hiddenFile = 0
With Fscann
 .tabson.AktifTab = 2
 .tabson.GantiJudul 3, "Virus @ 0"
 .tabson.GantiJudul 4, "Registry @ 0"
 .tabson.GantiJudul 5, "Hidden @ 0"
 .ucListView1(0).ListItems.Clear
 .ucListView1(1).ListItems.Clear
 .ucListView1(2).ListItems.Clear
 .progBSC.Value = 0
 .Command1(0).Enabled = False
  tmpCMD = .Command1(1).Caption
 .Command1(1).Caption = "Stop"

.Timer1.Enabled = True
End With
End Sub
Private Sub endSCAN()
 Fscann.Timer1_Timer
 Fscann.Timer1.Enabled = False
 Fscann.Command1(0).Enabled = True
 Fscann.Command1(1).Caption = tmpCMD
End Sub
Public Sub prosSCAN()
Dim ls As Collection
Dim i As Long, tmp As Long
With Fscann
Call prepScan
Set ls = New Collection
Fscann.DirTree1.OutPutPath ls
     If ProsesNode = True Then scanProC
     If StartUpNode = True Then scanStartUP
     'If RegNode = True Then ScanRegistry Fscann.Text1, Fscann.ucListView1(1), False
     
     For i = 1 To ls.Count
        Cari = False
        scanN ls(i), True
        tmp = Perc
        Perc = Perc + scannedF
        scannedF = tmp
        .Label2(2).Caption = Perc
        Cari = True
     scanN ls(i), True
     If scAn = False Then Exit For
     Next i
Set ls = Nothing
Call endSCAN
End With
End Sub
Public Sub ScanEXT(Folder As String, file As Boolean)
Dim tmp As Long
Load Fscann
Fscann.Show
Call prepScan
tunggON 5
If file = False Then
    Cari = False
    scanN Folder, True
    tmp = Perc
    Perc = Perc + scannedF
    scannedF = tmp
    Fscann.Label2(2).Caption = Perc
    Cari = True
    scanN Folder, True
Else
    LokasiD = Folder
    scannedF = 1
    Perc = 0 + scannedF
    scannedF = 1
    Fscann.Label2(2).Caption = Perc
    scanNOW Folder, "", False
    CheckAttrib Folder, False
End If
Call endSCAN
End Sub
Public Sub ShowProperties(filename As String, OwnerhWnd As Long)
On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = App.hInstance
        .lpIDList = 0
    End With
    ShellExecuteEx SEI
End Sub
Public Sub CekALL(Check1 As CheckBox, Lv As ucListView)
Dim i As Long
For i = 1 To Lv.ListItems.Count
If Check1.Value = 1 Then
If Lv.ListItems.Item(i).Checked = False Then
Lv.ListItems.Item(i).Checked = True
End If
Else
If Lv.ListItems.Item(i).Checked Then
Lv.ListItems.Item(i).Checked = False
End If
End If
DoEvents
Next i
End Sub
Public Sub shGantiwarna(sha As Shape)
If sha.BackColor = &HFF Then
sha.BackColor = &HFFFFFF
Else
sha.BackColor = &HFF
End If
End Sub
Public Sub FormCenter(frmObj As Form)
Dim lLeft As Long
Dim lTop As Long
On Error Resume Next
If frmObj.ScaleMode <> vbTwips Then frmObj.ScaleMode = vbTwips
lLeft = (Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN)) / 2
lLeft = lLeft - (frmObj.Width / 2)
lTop = (Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYFULLSCREEN)) / 2
lTop = lTop - (frmObj.Height / 2)
frmObj.Move lLeft, lTop
End Sub
Public Sub onTOP(lngHwnd As Long, Optional noonTop As Boolean)
If noonTop = False Then
SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
Else
SetWindowPos lngHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End If
End Sub
Public Sub SetFormPojok(Objek As Object)
Dim m_iScrnBottom As Long
Dim m_iOSver As Byte
Dim rc As RECT
Dim scrnRight As Long
Dim OSV As OSVERSIONINFO
OSV.OSVSize = Len(OSV)

If GetVersionEx(OSV) = 1 Then
    If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then m_iOSver = 1
    If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then m_iOSver = 2
End If

Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rc, 0&)
m_iScrnBottom = rc.Bottom * Screen.TwipsPerPixelY
scrnRight = (rc.Right * Screen.TwipsPerPixelX)
Objek.Move scrnRight - (Objek.Width + 100), m_iScrnBottom - (Objek.Height + 100), Objek.Width

SetWindowPos Objek.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub getProces(Lv As ucListView)
Dim jml As Long, a As Long
Dim lva As cListItem
Dim patZ As String
PamzEnumerateProcesses enumproC
Lv.ListItems.Clear
Set Lv.ImageList = Nothing
Set cImgList = New gComCtl
Set Lv.ImageList = cImgList.NewImageList(16, 16, imlColor32)

getICON
jml = UBound(enumproC)
For a = 0 To jml
patZ = PamzNtPathToUserFriendlyPathW(enumproC(a).szNtExecutablePathW)
If Len(patZ) > 1 Then
Set lva = Lv.ListItems.Add()
lva.Text = enumproC(a).szNtExecutableNameW
lva.IconIndex = a
lva.SubItem(2).Text = patZ
lva.SubItem(3).Text = enumproC(a).nProcessID
End If
Next a
Set lva = Nothing
Set cImgList = Nothing
End Sub
Sub getICON()
Dim i As Long, jml As Long
Dim patZ As String
jml = UBound(enumproC)
  For i = 0 To jml
    patZ = PamzNtPathToUserFriendlyPathW(enumproC(i).szNtExecutablePathW)
    DrawIco patZ, Fscann.picBuffer, ricnSmall
    Fscann.ucListView2(0).ImageList.AddFromDc Fscann.picBuffer.hdc, 16, 16
  Next i
End Sub
Public Sub getDBNFO(Lv As ucListView)
Dim i As Long, B As Long
Dim jVPE As Long, jVIC As Long
Dim nVPE() As String, nVIC() As String
Dim lva As cListItem

scE.InfoDB jVPE, nVPE(), jVIC, nVIC()
Lv.ListItems.Clear
For i = 0 To jVPE
B = B + 1
Set lva = Lv.ListItems.Add()
lva.Text = STR$(B)
lva.SubItem(2).Text = nVPE(i)
Next i

'For i = 0 To jVIC
'b = b + 1
'fMain.List1.AddItem "   " & b & ". " & nVIC(i)
'Next i
Set lva = Nothing
End Sub
Sub savesettingAPP()
Dim fileN As String
If isInstalled = True Then
fileN = fixP(ObAVdir) & "ObAV.cfg"
Else
fileN = fixP(App.path) & "ObAV.cfg"
End If
With Fscann
SetINI "Options", "chkop0", .chkOP(0).Value, fileN
SetINI "Options", "chkop1", .chkOP(1).Value, fileN
SetINI "Options", "chkop2", .chkOP(2).Value, fileN
SetINI "Options", "chkop3", .chkOP(3).Value, fileN
SetINI "Options", "chkop4", .chkOP(4).Value, fileN
SetINI "Options", "chkop5", .chkOP(5).Value, fileN
SetINI "Options", "chkop6", .chkOP(6).Value, fileN
SetINI "Options", "chkop7", .chkOP(7).Value, fileN
SetINI "Options", "chkop8", .chkOP(8).Value, fileN
SetINI "Options", "chkop9", .chkOP(9).Value, fileN
SetINI "Options", "Sound", .snd.Value, fileN
SetINI "Options", "Splash", .StArTup.Value, fileN
SetINI "Options", "ScanFD", .ScanFD.Value, fileN
End With
End Sub

Sub getsettingAPP()
On Error GoTo errHh
Dim i As Integer
Dim fileN As String
If isInstalled = True Then
fileN = fixP(ObAVdir) & "ObAV.cfg"
Else
fileN = fixP(App.path) & "ObAV.cfg"
End If
If scE.FileAdaX(fileN) = False Then GoTo errHh
With Fscann
.chkOP(0).Value = GetINI("Options", "chkop0", fileN)
.chkOP(1).Value = GetINI("Options", "chkop1", fileN)
.chkOP(2).Value = GetINI("Options", "chkop2", fileN)
.chkOP(3).Value = GetINI("Options", "chkop3", fileN)
.chkOP(4).Value = GetINI("Options", "chkop4", fileN)
.chkOP(5).Value = GetINI("Options", "chkop5", fileN)
.chkOP(6).Value = GetINI("Options", "chkop6", fileN)
.chkOP(7).Value = GetINI("Options", "chkop7", fileN)
.chkOP(8).Value = GetINI("Options", "chkop8", fileN)
.chkOP(9).Value = GetINI("Options", "chkop9", fileN)
.snd.Value = GetINI("Options", "Sound", fileN)
.StArTup.Value = GetINI("Options", "Splash", fileN)
.ScanFD.Value = GetINI("Options", "ScanFD", fileN)
For i = 0 To 10
.chkOP_Click i
Next i

End With
errHh:
End Sub
Public Sub cmdAktiv(rtpAKTIV As Boolean, sStatus As Shape, Label3 As Label, Command3 As CommandButton)
If rtpAKTIV = True Then
shGantiwarna sStatus
If Not Label3.Caption = "Active" Then _
Label3.Caption = "Active"
If Not Command3.Caption = "DeActivate" Then _
Command3.Caption = "DeActivate"
Else
If Not Label3.Caption = "Non Active" Then _
Label3.Caption = "Non Active"
If Not Command3.Caption = "Activate" Then _
Command3.Caption = "Activate"
End If
End Sub

