Attribute VB_Name = "bsScann"
Option Explicit
Private Const MAX_PATH  As Long = 260
Private Const MAX_BUF   As Long = 512

Private Const ojan_ARCHIVE = &H20
Private Const ojan_DEVICE = &H40
Private Const ojan_NORMAL = &H80
Private Const ojan_READONLY = &H1
Private Const ojan_HIDDEN = &H2
Private Const ojan_SYSTEM = &H4
Private Const ojan_DIRECTORY = &H10

Private Type ojanFILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type ojan_Cari_DATA
    dwFileAttributes    As Long
    ftCreationTime      As ojanFILETIME
    ftLastAccessTime    As ojanFILETIME
    ftLastWriteTime     As ojanFILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Private Declare Function ojanCariFileU Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function ojanCariLanjutU Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function ojanCariTutupU Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Boolean
Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal pv_lpFileName As Long) As Long

Public scE As aXlEngine
Public FgrdEnd As Boolean
Public loADingF As Boolean
Public Sub scanN(ByVal sPath As String, Optional ByVal FIRS As Boolean = False, _
Optional ByVal sCanALL As Boolean = True)
On Error Resume Next
Dim OCD As ojan_Cari_DATA
Dim szFullPath As String
Dim szFileName As String
Dim hFind As Long
Dim NextStack As Long
Dim zSlash As String
Dim ojanIsFolder As Boolean
Dim isHiden As Boolean
Dim ttk1 As String
Dim ttk2 As String

ttk1 = ChrW$(46)
ttk2 = ttk1 & ttk1
zSlash = ChrW$(42)
If FIRS = True Then scAn = True
If scAn = False Then GoTo erH
sPath = fixP(sPath)
    hFind = ojanCariFileU(StrPtr(sPath & zSlash), VarPtr(OCD))
    If hFind < 1 Then GoTo erH
Do
    If scAn = False Then Exit Do
    szFileName = TrimW(OCD.cFileName)
    szFullPath = sPath & szFileName
    ojanIsFolder = ((OCD.dwFileAttributes And ojan_DIRECTORY) = ojan_DIRECTORY)
    If szFileName <> ttk1 And szFileName <> ttk2 Then
lewaT:
    If pause = True Then
    Do: DoEvents
    If pause = False Then Exit Do
    Sleep 1: Loop
    End If
    If Cari = True Then
    LokasiD = szFullPath
    
    scanNOW sPath, szFileName, ojanIsFolder
    CheckAttrib szFullPath, ojanIsFolder
    
    End If
    
    scannedF = scannedF + 1
    DoEvents
    Else
    ojanIsFolder = False
    End If
    
    If ojanIsFolder = True Then
    If scAn = False Then Exit Do
    If sCanALL = True Then
    Call scanN(szFullPath, False)
    End If
    End If
    NextStack = ojanCariLanjutU(hFind, VarPtr(OCD))
Loop While NextStack
Call ojanCariTutupU(hFind)
szFullPath = vbNullString
szFileName = vbNullString
erH:
End Sub
Sub scanNOW(szDirektori As String, szNamaFile As String, bAdalahFolder As Boolean)
On Error GoTo aK
Dim pjG As Long
If bAdalahFolder = False Then
pjG = scE.lenFileEX(szDirektori & szNamaFile)
If pjG <= 2 Then Exit Sub
If pjG > 3750000 Then Exit Sub
    CekVirus szDirektori & szNamaFile, pjG
End If
aK:
pjG = vbNull
End Sub
Public Function CekVirus(FileNOW As String, pjG As Long, Optional pId As Long, _
Optional VBZ As Boolean = False) As Boolean

Dim ex As String, nmU As String, BineR As String
Dim suspected As Boolean

On Error GoTo erR
nmU = UCase$(FileNOW)
ex = Right$(nmU, 3)
BineR = scE.GFQ(FileNOW)
If cExVmx = True Then _
If scE.cEkVMX(ex) = True Then GoTo taIK

If VBA.Left$(BineR, 2) = "MZ" Then
    If cPEhead Then _
    If scE.CekDNA(StrConv(BineR, vbFromUnicode), pjG, ex) Then GoTo taIK
    If cVerHead Then _
    If scE.cekverHDR(FileNOW) Then GoTo taIK
    If cIcon Then _
    If scE.CekIconB(FileNOW, pjG) Then GoTo taIK
    If cString Then
        If scE.cekDocinF(BineR) Then GoTo taIK
        If scE.cekRunouce(BineR) Then GoTo taIK
        If scE.cekPirut(BineR) Then GoTo taIK
        If scE.cekAlman(BineR) Then GoTo taIK
        If scE.cekVirutGen(BineR) Then GoTo taIK
        If scE.cEkEICAR(BineR) Then GoTo taIK
        If pjG > 300000 Then GoTo erR
        If ex <> "EXE" Then GoTo erR
        If scE.cekTrojanGen(BineR) Then GoTo taIK
    End If
Else
    If ex = "LNK" And cSortcut = True Then _
        If scE.ceksorCuT(BineR) = True Then GoTo taIK
    If pjG > 75000 Then GoTo erR
    If VBZ = True Then GoTo ceKvBS
    If InStr("DB INI INF VBS DLS TXT EML BAT", ex) And cMalScrip = True Then
ceKvBS:
        If scE.CekFisiK(BineR) = True Then GoTo taIK
        If scE.cekRunouce(BineR) Then GoTo taIK
        If scE.cekCanTix(BineR) Then GoTo taIK
        If scE.cekunrealX(BineR) Then GoTo taIK
        If scE.cEkEICAR(BineR) Then GoTo taIK
    End If
    
End If

erR:
BineR = vbNullString
nmU = vbNullString
ex = vbNullString
Exit Function
taIK:
BineR = vbNullString
suspected = (tipeVirus = "suspect")
If suspected = False Then CekVirus = True
ex:
    If suspected = False Then
        If pId > 1 Then Killproc pId, True
    End If
    scaNERSC FileNOW
End Function
Sub scaNERSC(FileNOW As String)
Dim lis As cListItem
Set lis = Fscann.ucListView1(0).ListItems.Add(, viruzE)
lis.SubItem(2).Text = FileNOW
lis.SubItem(3).Text = tipeVirus
Set lis = Nothing
jmlVIR = jmlVIR + 1
Fscann.tabson.GantiJudul 3, "Virus @ " & jmlVIR
End Sub
Public Sub scanProC()
Dim c As Long, pjG As Long
List_Process
For c = 1 To jmlProcess

LokasiD = "memory => " & pPath(c)
If InStr(LCase(TrimW(pPath(c))), "wscript.exe") Then
scanProcScript ProcessId(c)
End If

pjG = scE.lenFileEX(pPath(c))
If pjG <= 2 Then GoTo lWT
If pjG > 1750000 Then GoTo lWT
CekVirus pPath(c), pjG, ProcessId(c)
lWT:
DoEvents
Next c
End Sub
Private Sub scanProcScript(pId As Long)
On Error GoTo erH
Dim pjG As Long
Dim RET As Collection
Dim objWMI As Object, objProc As Object, Item As Object
Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set objProc = objWMI.execquery("Select * from Win32_Process where ProcessID = """ & pId & """", , 48)
For Each Item In objProc
Set RET = Arguments(CStr(Item.CommandLine))
If RET.Count > 1 Then
    pjG = scE.lenFileEX(RET(RET.Count))
    If pjG <= 2 Then GoTo lWT
    If pjG > 1750000 Then GoTo lWT
    CekVirus RET(RET.Count), pjG, pId, True
lWT:
End If
Next Item
erH:
Set objWMI = Nothing
Set objProc = Nothing
Set Item = Nothing
End Sub
Public Sub CheckAttrib(sFile As String, bFolder As Boolean)
Dim NAT      As Long
Dim ObjName  As String
Dim sFileU As String
sFileU = UCase$(sFile)
If InStr(sFileU, "WINDOWS") Then Exit Sub
If InStr(sFileU, "DOCUME~1") Then Exit Sub
If InStr(sFileU, "DOCUMENTS AND SETT") Then Exit Sub
If InStr(sFile, "$$$obavQRN") Then Exit Sub
'If IsFileProtectedBySystem(sFile) = True Then Exit Sub

NAT = GetFileAttributesW(StrPtr(sFile))
ObjName = GetFileName(sFile)

If (NAT = 2 Or NAT = 34 Or NAT = 3 Or NAT = 6 Or NAT = 22 Or NAT = 18 Or NAT = 50 Or NAT = 19 Or NAT = 35) Then
    Fscann.ucListView1(2).ListItems.Add(, ObjName).SubItem(2).Text = sFile
    hiddenFile = hiddenFile + 1
    Fscann.tabson.GantiJudul 5, "Hidden @ " & hiddenFile
End If

End Sub
Public Sub getQuarantin(Lv As ucListView)
On Error Resume Next
Dim OCD As ojan_Cari_DATA
Dim szFileName As String
Dim hFind As Long
Dim NextStack As Long
Dim zSlash As String
Dim ojanIsFolder As Boolean
Dim ttk1 As String
Dim ttk2 As String
Dim sPath As String
Dim lva As cListItem
Dim oriPath As String
Dim FileNam As String
Dim jumlahisi As Long

sPath = dirC & dirKaRan
ttk1 = ChrW$(46)
ttk2 = ttk1 & ttk1
zSlash = ChrW$(42)
sPath = fixP(sPath)
Lv.ListItems.Clear

    hFind = ojanCariFileU(StrPtr(sPath & zSlash), VarPtr(OCD))
    If hFind < 1 Then GoTo erH
Do
    szFileName = TrimW(OCD.cFileName)
    FileNam = Left$(szFileName, Len(szFileName) - 4)
    oriPath = GetINI("quarantin", FileNam, dirC & fixP(dirKaRan) & "$$obav.dat")
    ojanIsFolder = ((OCD.dwFileAttributes And ojan_DIRECTORY) = ojan_DIRECTORY)
    If szFileName <> ttk1 And szFileName <> ttk2 And Right$(szFileName, 4) = "_vir" _
    And Len(oriPath) > 1 Then
        If ojanIsFolder = False Then
        Set lva = Lv.ListItems.Add()
        lva.Text = FileNam
        lva.SubItem(2).Text = sPath & szFileName
        lva.SubItem(3).Text = oriPath
        jumlahisi = jumlahisi + 1
        End If
    DoEvents
    End If

    NextStack = ojanCariLanjutU(hFind, VarPtr(OCD))
Loop While NextStack
Call ojanCariTutupU(hFind)
erH:
If jumlahisi = 0 Then
Kill dirC & fixP(dirKaRan) & "$$obav.dat"
End If
Set lva = Nothing
End Sub
