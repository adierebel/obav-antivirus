Attribute VB_Name = "ModUI"
'****************************************************
'Name       : fMakeATranpArea
'Author     : Sapta_Agunk
'Date       : 15-11-2010
'Purpose    : Create a Transprarent Area in a form so that you can see through
'Input      : Areatype : a String indicate what kind of hole shape it would like to make
'PCordinate : the cordinate area needed for create the shape:
'Example    : X1, Y1, X2, Y2 for Rectangle
'OutPut     : A boolean
'****************************************************
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_Defaut                As Long = &H2
Private Const WS_EX_LAYERED     As Long = &H80000
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Const LWA_ALPHA As Long = &H2
Private winHwnd As Long
Dim POS As Integer, StartPos As Integer, Lengh As Integer, iTeks As Integer
Dim MyTeks As String
Private noc As Integer
Private str1 As String
Private str2 As String
Private str3 As String
Private strmessage As String
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, _
ByVal nIndex As Long) As Long
Public Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Function fMakeATranspArea(AreaType As String, pCordinate() As Long) As Boolean
Const RGN_DIFF = 4
Dim lOriginalForm As Long
Dim ltheHole As Long
Dim lNewForm As Long
Dim lfWidth As Single
Dim lParam(1 To 6) As Long
Dim lfHeight As Single
Dim lborder_width As Single
Dim ltitle_height As Single
 On Error GoTo Trap
 lOriginalForm = CreateRectRgn(0, 0, lfWidth, lfHeight)
Select Case AreaType
 Case "Elliptic"
 ltheHole = CreateEllipticRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
 Case "RectAngle"
 ltheHole = CreateRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
 Case "RoundRect"
 ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(5), pCordinate(6))
 Case "Circle"
 ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(3), pCordinate(4))
 Case Else
 MsgBox "Unknown Shape!!"
 Exit Function
 End Select
 lNewForm = CreateRectRgn(0, 0, 0, 0)
 CombineRgn lNewForm, lOriginalForm, _
 ltheHole, RGN_DIFF
 fMakeATranspArea = True
Exit Function
Trap:
 MsgBox "error Occurred. Error # " & erR.Number & ", " & erR.Description
End Function
Public Function Transparency(ByVal hWnd As Long, Optional ByVal Col As Long = vbBlack, _
    Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
Dim DisplayStyle As Long
      If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
        DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
        Call SetWindowLong(hWnd, GWL_EXSTYLE, DisplayStyle)
    End If
    Transparency = (SetLayeredWindowAttributes(hWnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY)) <> 0)
    If Not erR.Number = 0 Then erR.Clear
End Function
Public Sub ActiveTransparency(M As Form, d As Boolean, F As Boolean, _
        T_Transparency As Integer, Optional Color As Long)
Dim B As Boolean
        If d And F Then
            B = Transparency(M.hWnd, Color, T_Transparency, False)
        ElseIf d Then
            B = Transparency(M.hWnd, 0, T_Transparency, True)
        Else
            B = Transparency(M.hWnd, , 255, True)
        End If
End Sub




