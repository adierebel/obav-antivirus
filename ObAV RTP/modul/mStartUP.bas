Attribute VB_Name = "bsStartUP"
Option Explicit
Dim argumen As String
Public Function FixStartP(sFile As String) As String
Dim sTmp        As String
Dim sSpecial    As String
Dim nNum        As Long
Dim iCount      As Long
Dim sWS As String, sWSA As String
If scE.FileAdaX(sFile) = False Then
    ' dapatkan awal dari drive:\
    If InStr(sFile, ":\") > 0 Then sTmp = Mid(sFile, InStr(sFile, ":\") - 1)
    sTmp = Replace(sTmp, Chr(34), "")
    If scE.FileAdaX(sTmp) = True Then GoTo KLIMAKS
    
        sWSA = Replace(sFile, Chr(34), "")
        nNum = InStr(sWSA, " ")
        If nNum > 0 Then
        sWS = Left$(sWSA, nNum - 1)
            If LCase(sWS) = "rundll32.exe" Then

            ElseIf LCase(sWS) = "wscript.exe" Then

            End If
        End If

    nNum = InStr(sFile, "/")
    If nNum > 0 Then
       sTmp = Left(sFile, nNum - 1)
    Else
       nNum = InStr(StrReverse(sFile), "-")
       If nNum > 0 Then
          sTmp = Left(sFile, Len(sFile) - nNum)
       Else
          sTmp = sFile
                nNum = InStr(sFile, " ")
                If nNum > 0 Then
                sTmp = Left(sFile, nNum - 1)
                End If
       End If
    End If
    
    Do
        If Right(sTmp, 1) = Chr(32) Then
            sTmp = Left(sTmp, Len(sTmp) - 1)
        Else
            sTmp = sTmp
        End If
    Loop While Right(sTmp, 1) = Chr(32)
    sTmp = Replace(sTmp, Chr(34), "")
    
    If InStr(sTmp, "\") = 0 Then
       sSpecial = GetSpecFolder(WINDOWS_DIR)
       If scE.FileAdaX(sSpecial & "\" & sTmp) = True Then
          sTmp = sSpecial & "\" & sTmp
       Else
          sSpecial = GetSpecFolder(SYSTEM_DIR) ' coba di system32 sekarang
          If scE.FileAdaX(sSpecial & "\" & sTmp) = True Then
             sTmp = sSpecial & "\" & sTmp
          End If
       End If
    End If
    If scE.FileAdaX(sTmp) = True Then sTmp = sTmp Else sTmp = "Access Denied"
Else
    sTmp = sFile
End If

KLIMAKS:
FixStartP = sTmp
End Function
