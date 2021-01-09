Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const IOCTL_REMOVE_NOTIFY = 2236420

Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
Private Declare Function RegCreateKeyUn Lib "advapi32.dll" Alias "RegCreateKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExUn Lib "advapi32.dll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValueUn Lib "advapi32.dll" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As Long) As Long
Private Declare Function RegOpenKeyUn Lib "advapi32.dll" Alias "RegOpenKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExUn Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const REG_SZ As Long = 1
Private Const REG_DWORD = 4
Public Const ERROR_SUCCESS As Long = 0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

'---------EXPORT------------------------
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare Function ShellExecuteW Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
'---------end EXPORT--------------------

Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal CSIDL As Long, ByVal fCreate As Boolean) As Boolean
Public Enum IDFolder
    ALL_USER_STARTUP = &H18
    WINDOWS_DIR = &H24
    SYSTEM_DIR = &H25
    PROGRAM_FILE = &H26
    USER_DOC = &H5
    USER_STARTUP = &H7
    RECENT_DOC = &H8
    DEKSTOP_PATH = &H19
    STAR_PROGRAMS = &H17
End Enum
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Dim ServState As SERVICE_STATE
Dim Installed As Boolean
Dim serviSstarted As Boolean

Public Function progbar(objek As PictureBox, geser As Long)
objek.Width = objek.Width + (geser * 40)
End Function
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
Public Function FileADA(sFile As String) As Boolean
Dim fso As Object
Set fso = CreateObject(StrReverse("tcejbometsyseliF.gnitpircS"))
    FileADA = fso.FileExists(sFile)
Set fso = Nothing
End Function
Public Function FolderADA(sFolder As String) As Boolean
Dim fso As Object
Set fso = CreateObject(StrReverse("tcejbometsyseliF.gnitpircS"))
    FolderADA = fso.FolderExists(sFolder)
Set fso = Nothing
End Function
Public Function isADMIN() As Boolean
    isADMIN = CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
End Function
Public Sub SaveString(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal strdata As String)
Dim keyhand As Long
strdata = BuffUni(strdata)
RegCreateKeyUn hKey, StrPtr(strPath), keyhand
RegSetValueExUn keyhand, StrPtr(strValue), 0, REG_SZ, ByVal StrPtr(strdata), Len(strdata)
RegCloseKey keyhand
End Sub
Public Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim RET As Long
    RegCreateKeyUn hKey, StrPtr(strPath), RET
    RegDeleteValueUn RET, StrPtr(strValue)
    RegCloseKey RET
End Sub
Private Function BuffUni(STR As String) As String
    BuffUni = STR & String$(Len(STR), ";")
End Function
Public Function fixP(ByVal paT As String) As String
On Error Resume Next
If Right$(paT, 1) <> ChrW$(92) Then
fixP = paT & ChrW$(92)
Else
fixP = paT
End If
End Function
Public Function ObAVdir() As String
ObAVdir = fixP(GetSpecFolder(PROGRAM_FILE)) & "ObAV"
End Function
Public Function isInstalled() As Boolean
If FileADA(fixP(ObAVdir) & "ObavSpk.sys") = False Then GoTo kozonG
If FileADA(fixP(ObAVdir) & "KprocMon.sys") = False Then GoTo kozonG
isInstalled = True
Exit Function
kozonG:
isInstalled = False
End Function
Public Function obavTMP() As String
On Error Resume Next
Dim buf As String
Dim fso As Object
Set fso = CreateObject(StrReverse("tcejbometsyseliF.gnitpircS"))
buf = fso.GetSpecialFolder(2)
obavTMP = fixP(buf) & "obavTemp"
If FolderADA(obavTMP) = False Then
MkDir obavTMP
End If
Set fso = Nothing
End Function
Public Function appFullpatH() As String
Dim buffeR As String
buffeR = Space$(255)
GetModuleFileNameW App.hInstance, StrPtr(buffeR), 255
appFullpatH = TrimW(buffeR)
End Function
Public Function TrimW(ByVal paT As String) As String
    On Error Resume Next
    Dim ksG As Long
    ksG = InStr(paT, ChrW$(0))
    If ksG > 0 Then
        TrimW = VBA.Left$(paT, (ksG - 1))
    Else
        TrimW = paT
    End If
End Function
Public Function GetSpecFolder(ByVal lpCSIDL As IDFolder) As String
Dim sPath As String
Dim lRet As Long
    sPath = String$(255, 0)
    lRet = SHGetSpecialFolderPath(0&, sPath, lpCSIDL, False)
    If lRet <> 0 Then
        GetSpecFolder = TrimW(sPath)
    End If
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
Public Sub ShellEXE(path As String, Optional argU As String)
ShellExecuteW 0&, StrPtr("open"), StrPtr(path), StrPtr(argU), StrPtr(vbNullString), 1
End Sub
Public Function SetINI(nApp As String, nKey As String, nVal As String, Filename As String)
On Error GoTo salah
WritePrivateProfileString nApp, nKey, nVal, Filename
Exit Function
salah:
End Function
Public Function BuildServisSPK(Optional silent As Boolean = False) As Boolean
    CheckService "ObavSpk"
    If Not Installed Then
        SetNTService "ObavSpk", "ObAV Kernel Process Killer", fixP(ObAVdir) & "ObavSpk.sys", SERVICE_KERNEL_DRIVER, SERVICE_SYSTEM_START, vbNullString
    End If
    StartNTService "ObavSpk"
End Function
Public Function ClirServisSPK() As Boolean
    CheckService "ObavSpk"
    If serviSstarted = True Then
        StopNTService "ObavSpk"
        Sleep 100
    End If
    DeleteNTService "ObavSpk"
End Function
Public Sub BKproc(Optional silent As Boolean = False)
    CheckService "KprocMon"
    If Not Installed Then
        SetNTService "KprocMon", "ObAV Kernel Process Monitoring", fixP(ObAVdir) & "KprocMon.sys", SERVICE_KERNEL_DRIVER, SERVICE_SYSTEM_START, vbNullString
    End If
    StartNTService "KprocMon"
End Sub
Public Sub delKproc()
    CheckService "KprocMon"
    If serviSstarted = True Then
        StopNTService "KprocMon"
        Sleep 100
    End If
    DeleteNTService "KprocMon"
End Sub
Public Function RemovNotiv(Optional silent As Boolean = False) As Boolean
Dim hFileSVC As Long
Dim dwBytesReturned As Long
hFileSVC = CreateFile("\\.\\KprocMon", GENERIC_READ + GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
If hFileSVC > 0 Then
    If DeviceIoControl(hFileSVC, IOCTL_REMOVE_NOTIFY, ByVal 0, 0, ByVal 0, 0, dwBytesReturned, ByVal 0) > 0 Then
    RemovNotiv = True
    Else
    If silent = False Then _
    MsgBox "Remov Notv XXX", vbCritical + vbSystemModal
    End If
    CloseHandle hFileSVC
Else
    If silent = False Then _
    MsgBox "gagal membuka device", vbCritical + vbSystemModal
End If
End Function
Private Sub CheckService(nama As String)
    If GetServiceConfig(nama) = 0 Then
        Installed = True

        ServState = GetServiceStatus(nama)
        Select Case ServState
            Case SERVICE_RUNNING
                serviSstarted = True
            Case SERVICE_STOPPED
                serviSstarted = False
            Case Else
                MsgBox "BUG", vbSystemModal, "ObAV Error"
        End Select
    Else
        Installed = False
        serviSstarted = False
    End If
End Sub
Public Sub tunggON(tim As Long)
Dim i As Long
For i = 0 To tim
Sleep 100
DoEvents
Next i
End Sub
