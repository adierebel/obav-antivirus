Attribute VB_Name = "bsArgument"
Option Explicit
Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function getcoMan(Optional nOmor As Long) As String
On Error GoTo erH
Dim retSTR As Long
Dim RET As Collection
Dim buff As String
retSTR = GetCommandLineW
buff = Space$(255)
CopyMemory ByVal StrPtr(buff), ByVal retSTR, 255
Set RET = Arguments(Trim$(TrimW(buff)))
If RET.Count > 1 Then
    If nOmor > 0 Then
    buff = RET(nOmor)
    Else
    buff = RET(RET.Count)
    End If
End If
getcoMan = buff
erH:
End Function
