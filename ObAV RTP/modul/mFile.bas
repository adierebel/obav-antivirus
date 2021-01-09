Attribute VB_Name = "bsFile"
Option Explicit
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal CSIDL As Long, ByVal fCreate As Boolean) As Boolean
Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function MoveFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long

Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
Private Const MOVEFILE_REPLACE_EXISTING = &H1

Public Enum IDFolder
    ALL_USER_STARTUP = &H18
    WINDOWS_DIR = &H24
    SYSTEM_DIR = &H25
    PROGRAM_FILE = &H26
    USER_DOC = &H5
    USER_STARTUP = &H7
    RECENT_DOC = &H8
    DEKSTOP_PATH = &H19
    STAR_PROGRAMS = &H17
End Enum
Private Const DocHeader = "ÐÏà¡±á"
Private Function obavTMP() As String
On Error Resume Next
Dim buf As String
Dim fso As Object
Set fso = CreateObject(StrReverse("tcejbometsyseliF.gnitpircS"))
buf = fso.GetSpecialFolder(2)
obavTMP = fixP(buf) & "obavTemp"
If FolderADA(obavTMP) = False Then
MkDir obavTMP
End If
Set fso = Nothing
End Function
Public Function GetFileVersi(paT As String) As Long
Dim ve As String
Dim a As VERHEADER
GetVerHeader paT, a
ve = a.FileVersion
ve = Replace(ve, ".", "")
GetFileVersi = ve
End Function
Public Function appFullpatH() As String
Dim Buffer As String
Buffer = Space$(255)
GetModuleFileNameW App.hInstance, StrPtr(Buffer), 255
appFullpatH = TrimW(Buffer)
End Function
Public Function pathME() As String
pathME = fixP(App.path)
End Function
Public Function ObAVdir() As String
ObAVdir = fixP(GetSpecFolder(PROGRAM_FILE)) & "ObAV"
End Function
Public Function isInstalled() As Boolean
If FileADA(fixP(ObAVdir) & "ObavSpk.sys") = False Then GoTo kozonG
If FileADA(fixP(ObAVdir) & "KprocMon.sys") = False Then GoTo kozonG
isInstalled = True
Exit Function
kozonG:
isInstalled = False
End Function
Public Function FileADA(sFile As String) As Boolean
Dim fso As Object
Set fso = CreateObject(StrReverse("tcejbometsyseliF.gnitpircS"))
    FileADA = fso.FileExists(sFile)
Set fso = Nothing
End Function
Public Function FolderADA(sFolder As String) As Boolean
Dim fso As Object
Set fso = CreateObject(StrReverse("tcejbometsyseliF.gnitpircS"))
    FolderADA = fso.FolderExists(sFolder)
Set fso = Nothing
End Function
Public Sub Eksekusi(ByVal Sandi As String, Lv As ucListView, Optional MainNN As Boolean = False)
On Error GoTo exitS
Dim patN As String
Dim i As Long, j As Long
Dim namaFil As String, newFil As String
j = Lv.ListItems.Count
If FolderADA(dirC & dirKaRan) = False Then
MkDir (dirC & dirKaRan)
SetFileAttributesW (StrPtr(dirC & dirKaRan)), vbHidden + vbSystem
End If

For i = 1 To j
If Lv.ListItems.Item(j).Checked = True Then

patN = Lv.ListItems.Item(j).SubItem(2).Text
SetFileAttributesW (StrPtr(patN)), vbNormal
If Lv.ListItems.Item(j).SubItem(3).Text = "suspect" Then
Lv.ListItems.Item(j).SubItem(4).Text = "Please submit to obAV LAB"
GoTo exitS
End If
If cleanDOC(patN) = False Then
        If Sandi = "hapus" Then
        DeleteFileW StrPtr(patN)
        Lv.ListItems.Remove j
        Else
        namaFil = Right$(patN, Len(patN) - InStrRev(patN, "\", , vbTextCompare))
        newFil = dirC & dirKaRan & "\" & namaFil & "_vir"
cekEXIS:
        If scE.FileAdaX(newFil) = True Then
        namaFil = "(" & acAk & ")" & namaFil
        newFil = dirC & dirKaRan & "\" & namaFil & "_vir": GoTo cekEXIS
        End If
        encripFile patN, newFil
        SetINI "quarantin", namaFil, patN, dirC & fixP(dirKaRan) & "$$obav.dat"
        Lv.ListItems.Remove j
        End If
End If
End If
exitS:
j = j - 1
If MainNN = True Then
'fMain.Lfile(1).Caption = j
End If
DoEvents
Next i
End Sub
Public Sub EksekusiFile(ByVal Sandi As String, fileNEM As String)
On Error GoTo exitS
Dim namaFil As String, newFil As String
If FolderADA(dirC & dirKaRan) = False Then
MkDir (dirC & dirKaRan)
SetFileAttributesW (StrPtr(dirC & dirKaRan)), vbHidden + vbSystem
End If

SetFileAttributesW (StrPtr(fileNEM)), vbNormal
If cleanDOC(fileNEM) = False Then
        If Sandi = "hapus" Then
        DeleteFileW StrPtr(fileNEM)
        Else
        namaFil = Right$(fileNEM, Len(fileNEM) - InStrRev(fileNEM, "\", , vbTextCompare))
        newFil = dirC & dirKaRan & "\" & namaFil & "_vir"
cekEXIS:
        If scE.FileAdaX(newFil) = True Then
        namaFil = "(" & acAk & ")" & namaFil
        newFil = dirC & dirKaRan & "\" & namaFil & "_vir": GoTo cekEXIS
        End If
        encripFile fileNEM, newFil
        SetINI "quarantin", namaFil, fileNEM, dirC & fixP(dirKaRan) & "$$obav.dat"
        End If
End If
exitS:
End Sub
Public Sub exsekusiHiden(Lv As ucListView)
Dim i As Long, j As Long
Dim patN As String
j = Lv.ListItems.Count
For i = 1 To j
If Lv.ListItems.Item(j).Checked = True Then
    patN = Lv.ListItems.Item(j).SubItem(2).Text
    If IsFileProtectedBySystem(patN) = False Then
    SetFileAttributesW (StrPtr(patN)), vbNormal
    Lv.ListItems.Remove j
    End If
End If
j = j - 1
DoEvents
Next i
End Sub
Private Function cleanDOC(sFileExe As String) As Boolean
On Error GoTo erhH
Static isiFile As String
isiFile = scE.GFQ(sFileExe)
If VBA.Left$(isiFile, 2) = "MZ" And InStr(isiFile, DocHeader) > 0 Then
    HealDoc sFileExe, sFileExe & ".doc"
    Kill sFileExe
    cleanDOC = True
Else
erhH:
    cleanDOC = False
End If
End Function

Private Function HealDoc(sFileExe As String, sTarget As String) As Boolean
On Error GoTo errR
Static iPointer As Long
Static isiFile As String

Open sFileExe For Binary As #2
    isiFile = Space(LOF(2))
    Get #2, , isiFile
Close #2

iPointer = InStr(isiFile, DocHeader)
isiFile = Mid$(isiFile, iPointer)
Open sTarget For Binary As #1
    Put #1, , isiFile
Close #1

HealDoc = True
errR:
End Function

Public Function GetSpecFolder(ByVal lpCSIDL As IDFolder) As String
Dim sPath As String
Dim lRet As Long
    sPath = String$(255, 0)
    lRet = SHGetSpecialFolderPath(0&, sPath, lpCSIDL, False)
    If lRet <> 0 Then
        GetSpecFolder = TrimW(sPath)
    End If
End Function

Public Function ValidFile(ByRef sFile As String) As Boolean
If PathIsDirectory(StrPtr(sFile)) = 0 And PathFileExists(StrPtr(sFile)) = 1 Then
    ValidFile = True
Else
    ValidFile = False
End If
End Function
