Attribute VB_Name = "bsEngine"
Option Explicit
Private Type IMAGEDOSHEADER
    e_magic As String * 2
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
  End Type
 
  Private Type IMAGE_SECTION_HEADER
    nameSec As String * 6
    PhisicalAddress As Integer
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
   
  End Type
 
  Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
  End Type
 
  Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    SIZE As Long
  End Type
 
 
  Private Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
  End Type
 
  Private Type IMAGE_NT_HEADERS
    Signature As String * 4
    FileHeader As IMAGE_FILE_HEADER
   OptionalHeader As IMAGE_OPTIONAL_HEADER
  End Type
 
  Private DOSHEADER As IMAGEDOSHEADER
  Private NTHEADER As IMAGE_NT_HEADERS
  Private SECTIONSHEADER() As IMAGE_SECTION_HEADER
Private Const nFoun As String = "Not Found"
Public Function PatEe() As String
PatEe = QualifyPath(App.path)
End Function
Public Function QualifyPath(sPath As String) As String
If Len(sPath) > 0 Then
If Right$(sPath, 1) <> "\" Then
QualifyPath = sPath & "\"
Else
QualifyPath = sPath
End If
Else
QualifyPath = ""
End If
End Function
Public Function GetVerHeader(ByVal fPN As String, ByRef oFP As VERHEADER)
Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim i As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(11) As String
Dim strTemp As String
Dim fPNw As Long
fPNw = StrPtr(fPN)

lngBufferlen = GetFileVersionInfoSizeW(fPNw, 0)
If lngBufferlen > 0 Then
 ReDim bytBuffer(lngBufferlen)
 lngRc = GetFileVersionInfoW(fPNw, 0&, lngBufferlen, bytBuffer(0))
 If lngRc <> 0 Then
 lngRc = VerQueryValueW(bytBuffer(0), StrPtr("\VarFileInfo\Translation"), lngVerPointer, lngBufferlen)
    If lngRc <> 0 Then
    MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
    lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
    strLangCharset = Hex$(lngHexNumber)
            
    Do While Len(strLangCharset) < 8
        strLangCharset = "0" & strLangCharset
    Loop
            
    strVersionInfo(0) = "CompanyName"
    strVersionInfo(1) = "FileDescription"
    strVersionInfo(2) = "FileVersion"
    
    For i = 0 To 2
     strBuffer = String$(255, 0)
     strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(i)
     lngRc = VerQueryValueW(bytBuffer(0), StrPtr(strTemp), lngVerPointer, lngBufferlen)
     If lngRc <> 0 Then
        lstrcpyW StrPtr(strBuffer), lngVerPointer
        strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
        strVersionInfo(i) = strBuffer
     Else
        strVersionInfo(i) = nFoun
     End If
    Next i
    End If
  End If
End If
    
    For i = 0 To 2
        If Len(strVersionInfo(i)) <= 1 Then strVersionInfo(i) = nFoun
    Next i
    
    oFP.CompanyName = strVersionInfo(0)
    oFP.FileDescription = strVersionInfo(1)
    oFP.FileVersion = strVersionInfo(2)
    
lngBufferlen = vbNull
lngDummy = vbNull
lngRc = vbNull
lngVerPointer = vbNull
lngHexNumber = vbNull
i = vbNull
Erase bytBuffer
Erase bytBuff
strBuffer = vbNullString
strLangCharset = vbNullString
Erase strVersionInfo
strTemp = vbNullString
fPNw = vbNull
End Function
Public Function CalcH(ByVal paT As String, ByVal biT As Long) As String
On Error GoTo erHAN
Dim Bin() As Byte
Dim ByteS As Long, i As Long
ReDim Bin(biT) As Byte
Open paT For Binary As #1
Get #1, , Bin
Close #1
For i = 0 To biT
ByteS = ByteS + Bin(i) ^ 2
Next i
CalcH = Hex$(ByteS)
erHAN:
End Function
Public Function CalcIcon(ByVal lpFileName As String, ByVal LbyTe As Long) As String
On Error GoTo erHAN
Dim PicPath As String
Dim ByteSum As String
Dim IconExist As Long
Dim hIcon As Long
Dim paT As Long
paT = StrPtr(lpFileName)
IconExist = ExtractIconEx(paT, 0, ByVal 0&, hIcon, 1)
If IconExist <= 0 Then
    IconExist = ExtractIconEx(paT, 0, hIcon, ByVal 0&, 1)
    If IconExist <= 0 Then Exit Function
End If

Fengine.Picture1.BackColor = vbWhite
DrawIconEx Fengine.Picture1.hdc, 0, 0, hIcon, 0, 0, 0, 0, &H3
DestroyIcon hIcon

PicPath = PatEe & "aa.ico"
SavePicture Fengine.Picture1.Image, PicPath

ByteSum = CalcH(PicPath, LbyTe)
DeleteFileW (StrPtr(PicPath))

CalcIcon = ByteSum
erHAN:
PicPath = vbNullString
ByteSum = vbNullString
IconExist = vbNull
hIcon = vbNull
paT = vbNull
End Function


Public Function ReadPE(ByRef DATA() As Byte) As Integer
  On Error GoTo ErrX
  Dim CNT As Long
  Dim u As Long
  Dim dumpSec() As Byte
  Dim jmlSect As Integer
  Dim BiggestSectionOff   As Long
  Dim SectionToSize       As Long

  CopyMemory DOSHEADER, DATA(CNT), Len(DOSHEADER)
  If DOSHEADER.e_magic <> "MZ" Then Exit Function
  CopyMemory NTHEADER, DATA(DOSHEADER.e_lfanew), Len(NTHEADER)
  CNT = CNT + DOSHEADER.e_lfanew + Len(NTHEADER)
  If NTHEADER.Signature <> "PE" & Chr(0) & Chr(0) Then Exit Function
  jmlSect = (NTHEADER.FileHeader.NumberOfSections - 1)
  ReDim SECTIONSHEADER(jmlSect)
  ReDim SecT(jmlSect)
  For u = 0 To jmlSect
  CopyMemory SECTIONSHEADER(u), DATA(CNT), Len(SECTIONSHEADER(0))
  CNT = CNT + Len(SECTIONSHEADER(0))
   FseC = Hex$(SECTIONSHEADER(u).VirtualSize)
   VsecX = SECTIONSHEADER(u).VirtualSize
   CseCX = Hex$(SECTIONSHEADER(u).Characteristics)
   NsecX = SECTIONSHEADER(u).nameSec
   SecT(u) = SECTIONSHEADER(u).nameSec
   
    If u > 0 Then
        If SECTIONSHEADER(u).PointerToRawData > BiggestSectionOff Then
             BiggestSectionOff = SECTIONSHEADER(u).PointerToRawData ' biasanya section terakhir
             SectionToSize = SECTIONSHEADER(u).SizeOfRawData
        End If
    Else
        BiggestSectionOff = SECTIONSHEADER(u).PointerToRawData ' awalnya baygkan terbesar ada yang pertama
        SectionToSize = SECTIONSHEADER(u).SizeOfRawData
    End If
  Next u
  nRealSizePE = BiggestSectionOff + SectionToSize
  
  dWptR = Hex$(NTHEADER.FileHeader.TimeDateStamp)
  dWptR = dWptR & Hex$(NTHEADER.FileHeader.Characteristics)
  dWptR = dWptR & Hex$(NTHEADER.OptionalHeader.AddressOfEntryPoint)
ReadPE = jmlSect
ErrX:
End Function
