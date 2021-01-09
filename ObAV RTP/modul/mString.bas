Attribute VB_Name = "bsString"
Option Explicit
Public RegNode      As Boolean
Public ProsesNode   As Boolean
Public StartUpNode  As Boolean
Public WinNode      As Boolean
Public DocNode      As Boolean
Public ProgNode     As Boolean
'-----------------------------
Public cString As Boolean
Public cPEhead As Boolean
Public cIcon As Boolean
Public cExVmx As Boolean
Public cVerHead As Boolean
Public cMalScrip As Boolean
Public cSortcut As Boolean
Public cAntidestroy As Boolean
Public cOntop As Boolean
Public cEnScanwith As Boolean
'-----------------------------
Public PathScan As String
Public LokasiD As String
Public jmlVIR As Long
Public Perc As Long
Public scannedF As Long
Public Cari As Boolean
Public pause As Boolean
Public scAn As Boolean
Public registriBad As Long
Public hiddenFile As Long

Private Const MagicByte As Byte = &HFF
Public Const dirKaRan As String = "$$$obavQRN"
Public Const verobAV As String = "ObAV AntiVirus 1.4 Final"
Public Declare Function PathIsDirectoryW Lib "shlwapi.dll" (ByVal pszPath As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function MoveFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long

Public scannedFATT As Long
Public CariATT As Boolean
Public PathScanATT As String
Public scAnATT As Boolean
Public Function TrimW(ByVal paT As String) As String
On Error Resume Next
Dim ksG As Long
ksG = InStr(paT, ChrW$(0))
If ksG > 0 Then
TrimW = VBA.Left$(paT, (ksG - 1))
Else
TrimW = paT
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
Public Function boLL(a As Boolean) As Long
If a = True Then
boLL = 1
Else
boLL = 0
End If
End Function
Public Function SetINI(nApp As String, nKey As String, nVal As String, FileName As String)
On Error GoTo salah
WritePrivateProfileString nApp, nKey, nVal, FileName
Exit Function
salah:
End Function
Public Function GetINI(nApp As String, nKey As String, FileName As String, Optional nDefault As String)
On Error GoTo salah
Dim RET As String
RET = String(255, 0)
GetPrivateProfileString nApp, nKey, nDefault, RET, 255, FileName
GetINI = VBA.Left$(RET, InStr(1, RET, Chr(0), vbTextCompare) - 1)
Exit Function
salah:
End Function
Public Function GetFileName(sFile As String) As String ' Mendapatkan nama file+extensi secara normal
On Error Resume Next
Dim tmp As String
Dim nTmp  As Long

    tmp = StrReverse(sFile)
    nTmp = InStr(tmp, "\")
    tmp = Left(tmp, nTmp - 1)

GetFileName = StrReverse(tmp)
End Function
Public Function Arguments(commandL As String) As Collection
On Error Resume Next
Dim cmd As String, cmdLen As Long, cmdLine As String
cmdLine = commandL
cmd = commandL
cmdLen = Len(cmd)
Dim RET As New Collection
Dim Append As Boolean
Dim InvertStart As Boolean
Dim CurPos As Long
Dim curChar As String

Append = True
Dim CurArg As String
For CurPos = 1 To cmdLen
    curChar = Mid$(cmdLine, CurPos, 1)
    Select Case curChar
        Case Space(1), vbTab
            Append = False
            If InvertStart Then Append = True
        Case Chr(34)
            InvertStart = Not InvertStart
            If InvertStart Then Append = True
        Case Else
            Append = True
    End Select

    Select Case Append
    Case True
        If curChar <> Chr(34) Then CurArg = CurArg & curChar
    Case False
        RET.Add CurArg
        CurArg = ""
    End Select
    DoEvents
Next
If CurArg <> "" Then RET.Add CurArg
Set Arguments = RET
End Function
Public Function dirC() As String
Dim buf As String
buf = Space(255)
GetWindowsDirectory buf, 255
buf = TrimW(buf)
dirC = Left$(buf, 3)
End Function
Public Function EnDEC(ByRef ByteArray() As Byte, Optional ByRef Password As String) As Byte()
  Dim PwdLen As Long
  Dim PwdAsc As Byte
  Dim i As Long
  Dim j As Long
  Dim LB As Long
  Dim UB As Long
  
  PwdLen = Len(Password)
  LB = LBound(ByteArray)
  UB = UBound(ByteArray)

    For j = 1 To PwdLen
      PwdAsc = Asc(Mid$(Password, j, 1)) Xor MagicByte
      For i = LB To UB Step PwdLen
        ByteArray(i) = ByteArray(i) Xor PwdAsc Xor (i And MagicByte)
      Next i
      LB = LB + 1
    Next j
  
    EnDEC = ByteArray()
End Function
Public Sub encripFile(file As String, newfile As String)
Dim buf() As Byte, Buf2() As Byte
buf = scE.GFQ(file, False)
If UBound(buf) <= 1 Then Exit Sub
Buf2 = EnDEC(buf(), "ojanblank")
Open dirC & "~obav.tmp" For Binary Access Write As #1
Put #1, , Buf2
Close #1
If scE.FileAdaX(newfile) = True Then DeleteFileW StrPtr(newfile)
If scE.FileAdaX(file) = True Then DeleteFileW StrPtr(file)
MoveFileW StrPtr(dirC & "~obav.tmp"), StrPtr(newfile)
Erase buf
Erase Buf2
End Sub
