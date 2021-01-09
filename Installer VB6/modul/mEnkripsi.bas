Attribute VB_Name = "mEnkripsi"
Option Explicit
Private Declare Function MoveFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Const OFS_MAXPATHNAME = 128
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_DELETE = &H4
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3
Private Const dirC As String = "C:\"
Private Const MagicByte As Byte = &HFF

Public Sub encripFile(ByRef sFile() As Byte, ByVal newfile As String)
Dim buf() As Byte, Buf2() As Byte
'buf = GFQ(file, False)
If UBound(sFile) <= 1 Then Exit Sub
Buf2 = EnDEC(sFile(), "kcsjdngue")
Open dirC & "~obavinstall.tmp" For Binary Access Write As #1
Put #1, , Buf2
Close #1
If FileAdaX(newfile) = True Then DeleteFileW StrPtr(newfile)
MoveFileW StrPtr(dirC & "~obavinstall.tmp"), StrPtr(newfile)
Erase buf
Erase Buf2
End Sub
Public Function FileAdaX(ByVal FilNAM As String) As Boolean
Dim heFile As Long
FileAdaX = False
 heFile = CreateFileW(StrPtr(FilNAM), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
 If heFile > 0 Then
 FileAdaX = True
 CloseHandle heFile
 End If
End Function
Public Function GFQ(strFilePath As String, Optional bolAsString = True)
  Dim arrFileMain() As Byte
  Dim lngSize As Long, lngRet As Long
  Dim lngFileHandle As Long
    lngFileHandle = CreateFileW(StrPtr(strFilePath), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
   If lngFileHandle > 0 Then
    lngSize = GetFileSize(lngFileHandle, 0)
    ReDim arrFileMain(lngSize) As Byte
    ReadFile lngFileHandle, arrFileMain(0), UBound(arrFileMain), lngRet, ByVal 0&
    CloseHandle lngFileHandle
    ReDim Preserve arrFileMain(UBound(arrFileMain) - 1)
    If bolAsString Then
        GFQ = StrConv(arrFileMain(), vbUnicode)
      Else
        GFQ = arrFileMain()
    End If
   End If
Erase arrFileMain
lngSize = vbNull
lngFileHandle = vbNull
lngRet = vbNull
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
