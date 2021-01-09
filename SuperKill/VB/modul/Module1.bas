Attribute VB_Name = "Module1"
Option Explicit
Private Const FILE_DEVICE_UNKNOWN = &H22
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const IOCTL_KILL_PROCCESS = FILE_DEVICE_UNKNOWN Or 14 Or &H800 Or 0
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
Private Const SERVICE_ERROR_IGNORE = &H0
Private Const SERVICE_CONTROL_STOP = &H1

Private Declare Function DeviceIoControl Lib "KERNEL32.DLL" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
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
Public Function BuildServis() As Boolean
Dim hSCManager As Long
Dim hService As Long
hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ALL_ACCESS)
If hSCManager > 0 Then
    hService = createservice(hSCManager, "ObavSpk", "ObavKernelProcessKiller", SERVICE_ALL, SERVICE_KERNEL_DRIVER, SERVICE_DEMAND_START, SERVICE_ERROR_IGNORE, fixP(App.Path) & "ObavSpk.sys", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    If hService > 0 Then
munggah:
        If StartService(hService, 0, 0) Then
        BuildServis = True
        Else
        MsgBox "gagal strat"
        End If
        CloseServiceHandle hService
    Else
        hService = OpenService(hSCManager, "ObavSpk", SERVICE_ALL)
        If hService > 0 Then GoTo munggah
        MsgBox "gagal membuat servis"
    End If
    CloseServiceHandle hSCManager
Else
    MsgBox "gagal scmgr"
End If
End Function
Public Function ClirServis() As Boolean
Dim hSCManager As Long
Dim hService As Long
Dim SS As SERVICE_STATUS
hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ALL_ACCESS)
If hSCManager > 0 Then
    hService = OpenService(hSCManager, "ObavSpk", SERVICE_ALL)
    If hService > 0 Then
    ControlService hService, SERVICE_CONTROL_STOP, SS
    DeleteService hService
    CloseServiceHandle hService
    End If
    CloseServiceHandle hSCManager
End If
End Function
Public Function KeKillProcess(ByVal pId As Long) As Boolean
Dim hFileSVC As Long
Dim drvInput As Long
Dim dwBytesReturned As Long
hFileSVC = CreateFile("\\.\\ojansuperkill", GENERIC_READ + GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
If hFileSVC > 0 Then
    drvInput = pId
    If DeviceIoControl(hFileSVC, IOCTL_KILL_PROCCESS, drvInput, 4, ByVal 0, 0, dwBytesReturned, ByVal 0) > 0 Then _
    KeKillProcess = True _
    Else _
    MsgBox "xxxxxxxxxxxxxx"
    CloseHandle hFileSVC
Else
    MsgBox "gagal membuka device" & hFileSVC
End If
End Function
Public Function fixP(ByVal paT As String) As String
On Error Resume Next
If Right$(paT, 1) <> ChrW$(92) Then
fixP = paT & ChrW$(92)
Else
fixP = paT
End If
End Function
