Attribute VB_Name = "bsTray"
Option Explicit
'Development by oj4nBL4NK (http://)
Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Public Declare Function GetFileVersionInfoSizeW Lib "version.dll" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long
Public Declare Function GetFileVersionInfoW Lib "version.dll" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function VerQueryValueW Lib "version.dll" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Any, puLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const TASKBARMESSAGE As String = "TaskbarCreated"
Public Const NOTIFYICONDATA_V1_SIZE As Long = 88
Public Const NOTIFYICONDATA_V2_SIZE As Long = 488
Public Const NOTIFYICONDATA_V3_SIZE As Long = 504
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_INFO = &H10
Public Const NIS_SHAREDICON = &H2
Public Const NOTIFYICON_VERSION = &H3
Public Const WM_USER As Long = &H400
Public Const WM_MYHOOK As Long = WM_USER + 1
Public Const WM_TIMER = &H113
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_LBUTTONDBLCLK As Long = &H203

Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Public Const NIM_ADD = &H0
Public Const NIM_SETVERSION = &H4
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const GWL_WNDPROC As Long = (-4)

Public Const APP_TIMER_EVENT_ID As Long = 998
Public Const APP_SYSTRAY_ID = 999
Public Const APP_TIMER_MILLISECONDS As Long = 15000
Public DefWindowProc As Long

Private tmrRunning As Boolean
Private NOTIFYICONDATA_SIZE As Long
Private c_lTm As Long
Private hIconT As Long
Private ToolTipT As String
Private Function IsShellVersion(ByVal version As Long) As Boolean
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   Const sDLLFile As String = "shell32.dll"

   nBufferSize = GetFileVersionInfoSizeW(StrPtr(sDLLFile), nUnused)
   If nBufferSize > 0 Then
      ReDim bBuffer(nBufferSize - 1) As Byte
      Call GetFileVersionInfoW(StrPtr(sDLLFile), 0&, nBufferSize, bBuffer(0))
      If VerQueryValueW(bBuffer(0), StrPtr("\"), lpBuffer, nUnused) = 1 Then
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
         IsShellVersion = nVerMajor >= version
      End If
   End If
End Function
Private Sub SetShellVersion()
   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE
   End Select
End Sub
Public Sub IconTooltip(hIcon As StdPicture, toolt As String)
hIconT = hIcon.handle
ToolTipT = toolt
c_lTm = RegisterWindowMessage(TASKBARMESSAGE)
End Sub
Public Function ShellTrayIconAdd(hwnd As Long, Optional modi As Boolean = False)
   Dim nid As NOTIFYICONDATA
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP Or NIF_INFO
      .dwState = NIS_SHAREDICON
      .hIcon = hIconT
      .szTip = ToolTipT & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      .uCallbackMessage = WM_MYHOOK
   End With
   
   If Shell_NotifyIcon(NIM_ADD, nid) = 1 Then
      Call Shell_NotifyIcon(NIM_SETVERSION, nid)
      If modi = False Then
      Call SubClass(hwnd)
      End If
      ShellTrayIconAdd = 1
   End If
       
End Function
Public Sub ShellTrayIconRemove(hwnd As Long)
   Dim nid As NOTIFYICONDATA
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
   End With
   If tmrRunning Then Call TimerStop(hwnd)
   Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub
Private Sub ShellTrayBalloonTipClose(hwnd As Long)
   Dim nid As NOTIFYICONDATA
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_TIP Or NIF_INFO
      .szTip = vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
   End With
   Call Shell_NotifyIcon(NIM_MODIFY, nid)
End Sub
Public Sub ShellTrayBalloonTipShow(hwnd As Long, nIconIndex As Long, sTitle As String, sMessage As String)
   Dim nid As NOTIFYICONDATA
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      .szInfoTitle = sTitle & vbNullChar
      .szInfo = sMessage & vbNullChar
   End With
   Call Shell_NotifyIcon(NIM_MODIFY, nid)
End Sub
Private Sub SubClass(hwnd As Long)
   On Error Resume Next
   DefWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnSubClass(hwnd As Long)
   If DefWindowProc <> 0 Then
      SetWindowLong hwnd, GWL_WNDPROC, DefWindowProc
      DefWindowProc = 0
   End If
End Sub
Private Sub TimerBegin(ByVal hwndOwner As Long, ByVal dwMilliseconds As Long)
   If Not tmrRunning Then
      If dwMilliseconds <> 0 Then
         tmrRunning = SetTimer(hwndOwner, APP_TIMER_EVENT_ID, dwMilliseconds, AddressOf TimerProc) = APP_TIMER_EVENT_ID
      End If
   End If
End Sub
Public Function TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
   Select Case uMsg
      Case WM_TIMER
         If idEvent = APP_TIMER_EVENT_ID Then
            If tmrRunning = True Then
               Call TimerStop(hwnd)
               Call ShellTrayBalloonTipClose(fRtpSystem.hwnd)
            End If
         End If
      Case Else
   End Select
End Function
Private Sub TimerStop(ByVal hwnd As Long)
   If tmrRunning = True Then
      Debug.Print "timer stopped"
      Call KillTimer(hwnd, APP_TIMER_EVENT_ID)
      tmrRunning = False
   End If
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   On Error Resume Next

   Select Case hwnd
      Case fRtpSystem.hwnd
         Select Case uMsg
            Case WM_MYHOOK
                Select Case lParam
                  Case WM_RBUTTONUP
                  Call SetForegroundWindow(fRtpSystem.hwnd)
                    fRtpSystem.PopupMenu fRtpSystem.mnu
                  Case WM_LBUTTONDBLCLK
                    ShellEXE fixP(ObAVdir) & "ObAV.exe"
                  Case NIN_BALLOONSHOW
                     Call TimerBegin(hwnd, APP_TIMER_MILLISECONDS)
                  Case NIN_BALLOONHIDE
                     Call TimerStop(hwnd)
                  Case NIN_BALLOONUSERCLICK
                     Call TimerStop(hwnd)
                  Case NIN_BALLOONTIMEOUT
                     Call TimerStop(hwnd)
               End Select
            Case Else
               WindowProc = CallWindowProc(DefWindowProc, hwnd, uMsg, wParam, lParam)
               'Exit Function
         End Select
      Case Else
      WindowProc = CallWindowProc(DefWindowProc, hwnd, uMsg, wParam, lParam)
   End Select
   
If uMsg = c_lTm Then
    ShellTrayIconAdd fRtpSystem.hwnd, True
End If

End Function
