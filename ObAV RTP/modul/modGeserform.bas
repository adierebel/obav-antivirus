Attribute VB_Name = "modGeserform"
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function ReleaseCapture Lib _
    "user32" () As Long
Private Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lparam As Any) As Long
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
    ByVal hDCDst As Long, _
    pptDst As Any, _
    psize As Any, _
    ByVal hDCSrc As Long, _
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

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_EXPLORER = &H80000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_HIDEREADONLY = &H4
Private Const LVM_FIRST = &H1000
Public Sub MoveForm(hwnd As Long)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub



