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
    hWnd As Long
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
Public Sub prosSCAN()
Dim ls As Collection
Dim i As Long, tmp As Long
With FrmScanner
Call prepScan
Set ls = New Collection
FrmScanner.DirTree1.OutPutPath ls
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
Private Sub endSCAN()
 FrmScanner.Timer2.Enabled = False
  FrmScanner.Timer1.Enabled = False
 FrmScanner.Command6.Caption = tmpCMD
End Sub
Private Sub prepScan()
scAn = True
jmlVIR = 0
scannedF = 0
Perc = 0
registriBad = 0
hiddenFile = 0
With FrmScanner
 .Label1(1).Caption = "Virus (0)"
 .Label1(2).Caption = "Registry (0)"
 .Label1(4).Caption = "Hidden (0)"
 .ucListView1(0).ListItems.Clear
 .ucListView1(1).ListItems.Clear
 .ucListView1(2).ListItems.Clear
 .progBSC.Value = 0
  tmpCMD = .Command6.Caption
 .Command6.Caption = "Stop"

.Timer1.Enabled = True
.Timer2.Enabled = True
End With
End Sub
Public Sub ScanEXT(Folder As String, file As Boolean)
Dim tmp As Long
Load FrmScanner
FrmScanner.Show
Call prepScan
tunggON 5
If file = False Then
    Cari = False
    scanN Folder, True
    tmp = Perc
    Perc = Perc + scannedF
    scannedF = tmp
    FrmScanner.Label2(2).Caption = Perc
    Cari = True
    scanN Folder, True
Else
    LokasiD = Folder
    scannedF = 1
    Perc = 0 + scannedF
    scannedF = 1
    FrmScanner.Label2(2).Caption = Perc
    scanN Folder, True, False
    CheckAttrib Folder, False
End If
Call endSCAN
End Sub
Public Sub ProcScanRTP(PatScan As String, Psubfolder As Boolean)
Dim tmp As Long
scAn = True
scannedF = 0
Perc = 0
Fpojok.Show
    Cari = False
    scanN PatScan, True, Psubfolder
    tmp = Perc
    Perc = Perc + scannedF
    scannedF = tmp
    Cari = True
    scanN PatScan, True, Psubfolder
    tunggON 5
Unload Fpojok
End Sub
Public Sub ShowProperties(filename As String, OwnerhWnd As Long)
On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = OwnerhWnd
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
m_iScrnBottom = rc.bottom * Screen.TwipsPerPixelY
scrnRight = (rc.Right * Screen.TwipsPerPixelX)
Objek.Move scrnRight - (Objek.Width + 100), m_iScrnBottom - (Objek.Height + 100), Objek.Width

SetWindowPos Objek.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
