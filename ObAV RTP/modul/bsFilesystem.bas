Attribute VB_Name = "bsFilesystem"
Option Explicit
Public DaftarFileSystem() As String ' menampung daftar file system
Private Declare Function SfcGetFiles Lib "sfcfiles.dll" (ByVal pv_lpFilesBuffer As Long, ByVal pv_lpFileCount As Long) As Long
Private Declare Function ExpandEnvironmentStringsW Lib "kernel32.dll" (ByVal pv_InputStrPath As Long, ByVal pv_OutputExpandedPath As Long, ByVal btCharBufferOutLen As Long) As Long

Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal pv_lpString As Long) As Long
Private Declare Function IsBadReadPtr Lib "kernel32.dll" (ByVal pv_lpMemAddress As Long, ByVal nLenInBytes As Long) As Long

Private Declare Sub RtlMoveMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal pSourceBuffer As Long, ByVal nBufferLengthToMove As Long) '<---sebenarnya namanya kurang sesuai, karena yang dilakukan adalah menyalin (copy) isi dari src ke dst.
Private Declare Sub RtlFillMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFill As Long, ByVal nByteNumber As Long) '<---harusnya byte,tapi memori 32 bit, jadi nggak apa-apa, asal tetap bernilai antara 0 sampai 255.
Private Declare Sub RtlZeroMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.

Private Type PROTECTED_FILES_STRUCT
    pFilePart                   As Long
    pFilePath                   As Long
    pFillial                    As Long
End Type

Private Type PROTECTED_FILES_OUTPUT
    szShortFilePath             As String
    szLongFilePath              As String
End Type

Private Function PamzGetProtectedFileList(ByRef PrFileList() As PROTECTED_FILES_OUTPUT) As Long
On Error Resume Next
    Erase PrFileList()
Dim nResultValue                As Long
Dim nRetValue                   As Long
Dim pProtFileAddr               As Long
Dim nProtFileCount              As Long
Dim PFS                         As PROTECTED_FILES_STRUCT
Dim DTurn                       As Long
Dim LEAX                        As Long
    nRetValue = SfcGetFiles(VarPtr(pProtFileAddr), VarPtr(nProtFileCount))
    If (pProtFileAddr > 0) And (nProtFileCount > 0) Then
        ReDim PrFileList(nProtFileCount - 1) As PROTECTED_FILES_OUTPUT
        For DTurn = 0 To (nProtFileCount - 1)
            Call RtlMoveMemory(VarPtr(PFS), pProtFileAddr + (DTurn * Len(PFS)), Len(PFS))
            LEAX = lstrlenW(PFS.pFilePath)
            If LEAX > 0 Then
                If IsBadReadPtr(PFS.pFilePath, LEAX) = 0 Then
                    With PrFileList(DTurn)
                        .szShortFilePath = String$(LEAX, 0)
                        Call RtlMoveMemory(StrPtr(.szShortFilePath), PFS.pFilePath, LEAX * 2) '---unichars ke bytes.
                        LEAX = ExpandEnvironmentStringsW(StrPtr(.szShortFilePath), 0, 0) '---cari tahu panjangnya dulu.
                        If LEAX > 0 Then
                            .szLongFilePath = String$(LEAX, 0)
                            LEAX = ExpandEnvironmentStringsW(StrPtr(.szShortFilePath), StrPtr(.szLongFilePath), LEAX) '---cari tahu nama panjangnya.
                        End If
                    End With
                End If
            End If
        Next
        nResultValue = nProtFileCount
    Else
        nResultValue = 0
    End If
LBL_BROADCAST_RESULT:
    PamzGetProtectedFileList = nResultValue
LBL_TERAKHIR:
    If erR.Number <> 0 Then
        erR.Clear
    End If
End Function


Public Function EnumFileSystem() As Long
Dim PFL()      As PROTECTED_FILES_OUTPUT
Dim LEAX       As Long
Dim CTurn      As Long
Dim DriveSys   As String
    
    LEAX = PamzGetProtectedFileList(PFL())
    If LEAX <= 0 Then
        EnumFileSystem = 0
        GoTo LBL_KELUAR
    Else
        ReDim DaftarFileSystem(LEAX + 9) As String
        DriveSys = Left(GetSpecFolder(WINDOWS_DIR), 3)
        For CTurn = 0 To (LEAX - 1) ' dari indek satu
            DaftarFileSystem(CTurn + 1) = PotongKarKananJelek(PFL(CTurn).szLongFilePath)
        Next
        
        ' Tambahan
        DaftarFileSystem(CTurn + 2) = DriveSys & "ntldr"
        DaftarFileSystem(CTurn + 3) = DriveSys & "boot.ini"
        DaftarFileSystem(CTurn + 4) = DriveSys & "pagefile.sys"
        DaftarFileSystem(CTurn + 5) = DriveSys & "NTDETECT.COM"
        
        'cadangan (siapa tahu OS nya 2)
        DriveSys = "C:\"
        DaftarFileSystem(CTurn + 6) = DriveSys & "ntldr"
        DaftarFileSystem(CTurn + 7) = DriveSys & "boot.ini"
        DaftarFileSystem(CTurn + 8) = DriveSys & "pagefile.sys"
        DaftarFileSystem(CTurn + 9) = DriveSys & "NTDETECT.COM"
        
    End If
    Erase PFL()
    
    EnumFileSystem = LEAX + 7

Exit Function
LBL_KELUAR:
End Function

' EnumFileSystem harus dipanggi dulu (satu kali ajh)
Public Function IsFileProtectedBySystem(sFile As String) As Boolean
Dim iCounter As Long
On Error GoTo LBL_FALSE
For iCounter = 1 To UBound(DaftarFileSystem)
    If UCase(sFile) = UCase(DaftarFileSystem(iCounter)) Then
       IsFileProtectedBySystem = True
       Exit Function
    End If
    DoEvents
Next

LBL_FALSE:
IsFileProtectedBySystem = False
End Function

Private Function PotongKarKananJelek(sKar As String)
Dim LenKar As String
LenKar = Len(sKar)
If Asc(Right(sKar, 1)) < 20 Then
   PotongKarKananJelek = Mid(sKar, 1, LenKar - 1)
Else
   PotongKarKananJelek = sKar
End If
End Function

