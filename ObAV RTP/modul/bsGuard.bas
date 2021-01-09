Attribute VB_Name = "bsGuard"
Option Explicit
Public Sub cmdAktiv(rtpAKTIV As Boolean, sStatus As Shape, Label3 As Label, Command3 As CommandButton)
If rtpAKTIV = True Then
shGantiwarna sStatus
If Not Label3.Caption = "Active" Then _
Label3.Caption = "Active"
If Not Command3.Caption = "DeActivate" Then _
Command3.Caption = "DeActivate"
Else
If Not Label3.Caption = "Non Active" Then _
Label3.Caption = "Non Active"
If Not Command3.Caption = "Activate" Then _
Command3.Caption = "Activate"
End If
End Sub
