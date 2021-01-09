Attribute VB_Name = "modGeserform"
Private Declare Function ReleaseCapture Lib _
    "user32" () As Long
Private Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Function SetWindowPos Lib _
    "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib _
    "user32" (ByVal hwnd As Long, _
    ByVal hdcDst As Long, _
    pptDst As Any, _
    psize As Any, _
    ByVal hdcSrc As Long, _
    pptSrc As Any, _
    crKey As Long, _
    ByVal pblend As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib _
    "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib _
    "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Private Declare Function BeepAPI Lib _
    "kernel32" Alias "Beep" (ByVal dwFreq As Long, _
    ByVal dwDuration As Long) As Long
Private Declare Function GetSaveFileName Lib _
    "comdlg32.dll" Alias "GetSaveFileNameA" ( _
    lpofn As OPENFILENAME) As Long
Private Declare Function OpenProcessToken Lib _
    "advapi32" (ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib _
    "kernel32" () As Long
Private Declare Function ExitWindowsEx Lib _
    "user32" (ByVal uFlags As Long, _
    ByVal dwReserved As Long) As Long

Public Sub MoveForm(hwnd As Long)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub


