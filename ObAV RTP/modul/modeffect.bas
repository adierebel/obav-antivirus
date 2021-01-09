Attribute VB_Name = "modeffect"
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long

Private Declare Function UpdateLayeredWindow Lib "user32" _
(ByVal hwnd As Long, ByVal hDCDst As Long, pptDst As Any, _
psize As Any, ByVal hDCSrc As Long, pptSrc As Any, crKey As Long, _
ByVal pblend As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
MakeTransparent = 1
Else
Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
Msg = Msg Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, Msg
SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
MakeTransparent = 0
End If
If erR Then
MakeTransparent = 2
End If
End Function




