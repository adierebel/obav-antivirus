Attribute VB_Name = "bsDriver"
Option Explicit
Private Const FILE_DEVICE_UNKNOWN = &H22
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const IOCTL_KILL_PROCCESS = FILE_DEVICE_UNKNOWN Or 14 Or &H800 Or 0
Private Const IOCTL_CALL_BEEPER = FILE_DEVICE_UNKNOWN Or 14 Or &H800 Or 0
Private Const IOCTL_SET_NOTIFY = 2269184
Private Const IOCTL_REMOVE_NOTIFY = 2236420
Private Const IOCTL_GET_PROCESS_DATA = 2252808

Private Const INFINITE = -1

Private Type PROCESS_DATA
    bCreate         As Long
    dwProcessId     As Long
End Type

Private Type SERVICE_STATUS
        dwServiceType As Long
        dwCurrentState As Long
        dwControlsAccepted As Long
        dwWin32ExitCode As Long
        dwServiceSpecificExitCode As Long
        dwCheckPoint As Long
        dwWaitHint As Long
End Type
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SC_MANAGER_CONNECT = &H1
Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Private Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE)

Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4&
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_ALL = (STANDARD_RIGHTS_REQUIRED Or SERVICE_START Or SERVICE_STOP Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_INTERROGATE)

Private Const SERVICE_KERNEL_DRIVER = &H1
Private Const SERVICE_DEMAND_START = &H3
Private Const SERVICE_SYSTEM_START = &H1
Private Const SERVICE_ERROR_IGNORE = &H0
Private Const SERVICE_CONTROL_STOP = &H1

Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function createservice Lib "advapi32" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long

Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long

Public ghEvent As Long
Public exiTloop As Boolean
Public KprocID As Long

'==============================SERVICE CONTROLS=======================================
Public Function cretSVC(SvName As String, DisplayName As String, path As String, Optional silent As Boolean = False) As Boolean
Dim hSCManager As Long
Dim hService As Long
hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ALL_ACCESS)
If hSCManager > 0 Then
    hService = createservice(hSCManager, SvName, DisplayName, SERVICE_ALL, SERVICE_KERNEL_DRIVER, SERVICE_SYSTEM_START, SERVICE_ERROR_IGNORE, path, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    If hService > 0 Then
munggaH:
        If StartService(hService, 0, 0) Then
        cretSVC = True
        Else
        If silent = False Then _
        MsgBox "Failed Starting Driver", vbCritical + vbSystemModal
        End If
        CloseServiceHandle hService
    Else
        hService = OpenService(hSCManager, SvName, SERVICE_ALL)
        If hService > 0 Then GoTo munggaH
        If silent = False Then _
        MsgBox "Failed Creating Service" & " " & SvName, vbCritical + vbSystemModal
    End If
    CloseServiceHandle hSCManager
Else
    If silent = False Then _
    MsgBox "Failed Opening ServiceManager", vbCritical + vbSystemModal
End If
End Function
Public Function DelSVC(SvName As String)
Dim hSCManager As Long
Dim hService As Long
Dim SS As SERVICE_STATUS
hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ALL_ACCESS)
If hSCManager > 0 Then
    hService = OpenService(hSCManager, SvName, SERVICE_ALL)
    If hService > 0 Then
    ControlService hService, SERVICE_CONTROL_STOP, SS
    Sleep 100
    DeleteService hService
    CloseServiceHandle hService
    End If
    CloseServiceHandle hSCManager
End If
End Function
'=============================END SERVICE CONTROLS========================================

Public Function BuildServisSPK(Optional silent As Boolean = False) As Boolean
cretSVC "ObavSpk", "ObavKernelProcessKiller", fixP(ObAVdir) & "ObavSpk.sys", silent
End Function
Public Function ClirServisSPK() As Boolean
DelSVC "ObavSpk"
End Function
Public Function KeKillProcess(ByVal pId As Long, Optional silent As Boolean = False) As Boolean
Dim hFileSVC As Long
Dim dwBytesReturned As Long
hFileSVC = CreateFile("\\.\\ojansuperkill", GENERIC_READ + GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
If hFileSVC > 0 Then
    If DeviceIoControl(hFileSVC, IOCTL_KILL_PROCCESS, pId, 4, ByVal 0, 0, dwBytesReturned, ByVal 0) > 0 Then
    KeKillProcess = True
    Else
    If silent = False Then _
    MsgBox "Unable to Control Device Driver", vbCritical + vbSystemModal
    End If
    CloseHandle hFileSVC
Else
    If silent = False Then _
    MsgBox "Unable to Open Device Driver", vbCritical + vbSystemModal
End If
End Function
'=========================KprocMON=========================
Public Sub BKproc(Optional silent As Boolean = False)
cretSVC "KprocMon", "ObavKernelProcessMonitor", fixP(ObAVdir) & "KprocMon.sys", silent
End Sub
Public Sub delKproc()
DelSVC "KprocMon"
End Sub
Public Function SetNotiv(Optional silent As Boolean = False) As Boolean
Dim hFileSVC As Long
Dim Tredaidi As Long
Dim dwBytesReturned As Long
hFileSVC = CreateFile("\\.\\KprocMon", GENERIC_READ + GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
If hFileSVC > 0 Then
    ghEvent = CreateEvent(ByVal 0, False, False, vbNullString)
    exiTloop = False
    ojanApi.CreateThread 0, 0, AddressOf kepetProc, 0, 0, Tredaidi
    If DeviceIoControl(hFileSVC, IOCTL_SET_NOTIFY, ghEvent, 4, ByVal 0, 0, dwBytesReturned, ByVal 0) > 0 Then
    SetNotiv = True
    Else
    If silent = False Then _
    MsgBox "Gagal MengSet KprocMon", vbCritical + vbSystemModal
        If DeviceIoControl(hFileSVC, IOCTL_REMOVE_NOTIFY, ByVal 0, 0, ByVal 0, 0, dwBytesReturned, ByVal 0) > 0 Then
        DeviceIoControl hFileSVC, IOCTL_SET_NOTIFY, ghEvent, 4, ByVal 0, 0, dwBytesReturned, ByVal 0
        End If
    End If
    CloseHandle hFileSVC
Else
    If silent = False Then _
    MsgBox "gagal Opening KprocMon", vbCritical + vbSystemModal
End If
End Function
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
    exiTloop = True
    SetEvent ghEvent
    Sleep 100
    CloseHandle ghEvent
    CloseHandle hFileSVC
Else
    If silent = False Then _
    MsgBox "gagal membuka device", vbCritical + vbSystemModal
End If
End Function

'============================MULTI THREADING======================
Public Sub kepetProc(arg As Long)
Dim hwndSvr As Long
Dim hFileSVC As Long
Dim dwBytesReturned As Long
Dim DATA_PROC As PROCESS_DATA
munggaH:
If ojanApi.WaitForSingleObjectX(ghEvent, INFINITE) = INFINITE Then GoTo Lompat1
If exiTloop = True Then GoTo exitLOOPING
hFileSVC = ojanApi.CretFileX("\\.\\KprocMon", GENERIC_READ + GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
If hFileSVC > 0 Then
    If ojanApi.DevisIOCTLXX(hFileSVC, IOCTL_GET_PROCESS_DATA, 0, 0, VarPtr(DATA_PROC), LenB(DATA_PROC), dwBytesReturned, 0) > 0 Then
        If DATA_PROC.bCreate = 1 Then
        hwndSvr = ojanApi.FindWindow("#32770", "obav_22091993")
        KprocID = DATA_PROC.dwProcessId
        ojanApi.PostMeseg hwndSvr, WM_USER + 878&, DATA_PROC.dwProcessId, 0
        End If
    Else
    ojanApi.mesBOX 0, "GAGALprocesscaled", "x", 0
    End If
    ojanApi.CloseHandle hFileSVC
End If
Lompat1:
GoTo munggaH
exitLOOPING:
End Sub
