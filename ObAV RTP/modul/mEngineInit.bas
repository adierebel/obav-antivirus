Attribute VB_Name = "bSEngineInit"
Option Explicit

Public Type VERHEADER
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    Comments As String
    LegalTradeMarks As String
    PrivateBuild As String
    SpecialBuild As String
End Type

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(256) As Byte
End Type

Public Const OFS_MAXPATHNAME = 128
Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_DELETE = &H4
Public Const OPEN_EXISTING = 3

'--------icon--------------------------
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExW" (ByVal lpszFile As Long, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
'--------end icon----------------------

'---------manipulasi file--------------
Public Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Any, ByVal Length As Long)
Public Declare Function GetFileVersionInfoW Lib "version.dll" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSizeW Lib "version.dll" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long
Public Declare Function VerQueryValueW Lib "version.dll" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Any, puLen As Long) As Long
Public Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hHandle As Long) As Long
'---------end manipulasi file-----------

'---------EXPORT------------------------
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
'---------end EXPORT--------------------

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public nRealSizePE As Long
Public dWptR As String
Public SecT() As String
Public FseC As String

Public CseCX As String
Public NsecX As String
Public VsecX As String
 
Public jmlIco As Long
Public viruzE As String
Public tipeVirus As String
Public rtpPATH As axMemory
Public rtpSTATE As axMemory
Public guardPATH As axMemory
Public guardSTATE As axMemory
Public quiteSTATE As axMemory
Private Sub Main()
'MsgBox getcoMan(4)
'Exit Sub
'--------------------------------
Call initializeALL

Load FrmScanner
FrmScanner.Show

'If getcoMan = "-rtp" Then
    'If isInstalled = True Then
    'Fsplash.starTT
    'End If
'End If
End Sub
Private Sub initializeALL()
Call loadDB
Call EnumFileSystem
App.Title = vbNullChar
RamnitSrc = "R" & Chr(0) & "E" & Chr(0) & "C" & Chr(0) & "Y" & Chr(0) & "C" & Chr(0) & "L" & Chr(0) & "E" & Chr(0) & "R"
REGrun = "C" & Chr(0) & "u" & Chr(0) & "r" & Chr(0) & "r" & Chr(0) & "e" & Chr(0) & "n" & Chr(0) & "t" & Chr(0) & "V" & Chr(0) & "e" & Chr(0) & "r" & Chr(0) & "s" & Chr(0) & "i" & Chr(0) & "o" & Chr(0) & "n" & Chr(0) & "\" & Chr(0) & "R" & Chr(0) & "u" & Chr(0) & "n"
REGhiden = "C" & Chr(0) & "u" & Chr(0) & "r" & Chr(0) & "r" & Chr(0) & "e" & Chr(0) & "n" & Chr(0) & "t" & Chr(0) & "V" & Chr(0) & "e" & Chr(0) & "r" & Chr(0) & "s" & Chr(0) & "i" & Chr(0) & "o" & Chr(0) & "n" & Chr(0) & "\" & Chr(0) & "E" & Chr(0) & "x" & Chr(0) & "p" & Chr(0) & "l" & Chr(0) & "o" & Chr(0) & "r" & Chr(0) & "e" & Chr(0) & "r" & Chr(0) & "\" & Chr(0) & "A" & Chr(0) & "d" & Chr(0) & "v" & Chr(0) & "a" & Chr(0) & "n" & Chr(0) & "c" & Chr(0) & "e" & Chr(0) & "d"
Set scE = New aXlEngine
Set rtpPATH = New axMemory
Set rtpSTATE = New axMemory
Set guardPATH = New axMemory
Set guardSTATE = New axMemory
Set quiteSTATE = New axMemory

rtpPATH.OpenMemory "obav_rtppath"
rtpSTATE.OpenMemory "obav_rtpstate"
guardPATH.OpenMemory "obav_guardpath"
guardSTATE.OpenMemory "obav_guardstate"
quiteSTATE.OpenMemory "obav_quitestate"
cString = True
cPEhead = True
cIcon = True
cExVmx = True
cVerHead = True
cMalScrip = True
cSortcut = True
cAntidestroy = True
End Sub
Private Sub loadDB()
Dim i As Long
ViriiconNa = Split(IntViriVariantName, "|")
ViriIconID = Split(IntViriIconID, "|")
VirPEid = Split(virPE, "|")
virPEname = Split(virPEN, "|")
VirVERid = Split(VirVER, "|")
virVERname = Split(VirVERN, "|")

jmlIco = UBound(ViriIconID)
jmlVPE = UBound(VirPEid)
jmlVER = UBound(VirVERid)

erR:
i = vbNull
End Sub
Public Function acAk() As String
    Dim sTitle() As Variant
    sTitle = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0")
    Randomize
    acAk = sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & _
    sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle))
End Function
Public Sub RegObavExt(reg As Boolean)
Dim hMod As Long, pReg As Long
hMod = LoadLibrary(fixP(ObAVdir) & "ObavExt.dll")
If reg = True Then
pReg = GetProcAddress(hMod, "DllRegisterServer")
Else
pReg = GetProcAddress(hMod, "DllUnregisterServer")
End If
If pReg > 0 Then
CallWindowProc pReg, 0, ByVal 0&, ByVal 0&, ByVal 0&
End If
End Sub
