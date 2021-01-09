Attribute VB_Name = "mMAIN"
Option Explicit
Public Const INFINITE = -1&
Private Const WAIT_TIMEOUT = 258&
Private Const msgSETFG = 4160

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type
Public Const VER_PLATFORM_WIN32_NT = 2&
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent
Public IsNTService As Boolean
Public ServiceName() As Byte, ServiceNamePtr As Long
Dim ServState As SERVICE_STATE
Dim Installed As Boolean
Dim serviSstarted As Boolean

Private Sub MainSvc()
    Dim hnd As Long
    Dim h(0 To 1) As Long

    hStopEvent = CreateEvent(0, 1, 0, vbNullString)
    hStopPendingEvent = CreateEvent(0, 1, 0, vbNullString)
    hStartEvent = CreateEvent(0, 1, 0, vbNullString)
    ServiceName = StrConv(Service_Name, vbFromUnicode)
    ServiceNamePtr = VarPtr(ServiceName(LBound(ServiceName)))

        hnd = StartAsService
        h(0) = hnd
        h(1) = hStartEvent
        IsNTService = WaitForMultipleObjects(2&, h(0), 0&, INFINITE) = 1&
        If Not IsNTService Then
            CloseHandle hnd
            MessageBox 0&, "This program must be started as a service.", App.Title, msgSETFG
        End If
    
    If IsNTService Then
        SetServiceState SERVICE_RUNNING
        Do: DoEvents
        'tambah disini
        If ProsesadA(Medir & "ObavGuard.exe") = False Then
            Call_ObAV
        End If
        Loop While WaitForSingleObject(hStopPendingEvent, 2000&) = WAIT_TIMEOUT
        
        SetServiceState SERVICE_STOPPED
        SetEvent hStopEvent
        WaitForSingleObject hnd, INFINITE
        CloseHandle hnd
    End If
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
End
End Sub
Public Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT
End Function

Sub Main()
If Not CheckIsNT() Then
    MsgBox "Run this program on NT system"
    End
    Exit Sub
End If

AppPath = App.Path
If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
Select Case Trim$(Command$)
    Case "-svc"
        Call MainSvc
    Case "-install"
        Call CheckService
        If Not Installed Then
          SetNTService
        End If
    Case "-uninstall"
        Call CheckService
        If Installed = True And serviSstarted = False Then
          DeleteNTService
        End If
    Case "-start"
        Call CheckService
        If ServState = SERVICE_STOPPED And Installed = True Then
        StartNTService
        End If
    Case "-stop"
        Call CheckService
        If ServState = SERVICE_RUNNING Then
        StopNTService
        End If
    Case Else
        ExitProcess 0
End Select
End Sub

Private Sub CheckService()
    If GetServiceConfig() = 0 Then
        Installed = True

        ServState = GetServiceStatus()
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
