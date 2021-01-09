VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aXlEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const FILE_SHARE_READ = &H1
Public Function cEkVMX(ex As String) As Boolean
If ex = "VMX" Then
viruzE = "Conficker.AllVersion"
tipeVirus = "Worm"
cEkVMX = True
End If
End Function
Public Function CekDNA(Bin() As Byte, LbyTe As Long, exT As String) As Boolean
Dim h As String, g As String, jmlSEKs As Long, i As Long
Dim SecAkhir As String

jmlSEKs = ReadPE(Bin)

'If exT = "EXE" Then
'If LbyTe > (nRealSizePE + 20000) Then
'    viruzE = "Heur/pe.infected"
'    tipeVirus = "suspect"
'    CekDNA = True: GoTo r
'End If
'End If

'--------cek virut infection-------------
SecAkhir = Replace(NsecX, ".", vbNullString)
If Len(SecAkhir) < 6 Then GoTo lompat_1
If VsecX <> "4096" Then GoTo lompat_1
If CseCX <> "C0000000" Then GoTo lompat_1
        viruzE = "Infected Virut.56"
        tipeVirus = "Heuristic"
        CekDNA = True: GoTo r
lompat_1:
'MsgBox Len(SecAkhir) & " - " & VsecX & " - " & CseCX
'----------------------------------------

    If PEtemp = dWptR Then
        viruzE = "ObAV Heuristic.v2"
        tipeVirus = "Heuristic"
        CekDNA = True: GoTo r
    End If
    If InStrRev(virPE, dWptR) <> 0 Then
    For i = 0 To jmlVPE
        If VirPEid(i) = dWptR Then
        tipeVirus = "Worm"
        viruzE = virPEname(i)
        CekDNA = True: GoTo r
        End If
    Next i
    End If
If jmlSEKs > 2 Then

h = SecT(1)
g = SecT(jmlSEKs)
h = Replace(h, ".", vbNullString)
g = Replace(g, ".", vbNullString)
h = VBA.Left$(h, 4)

If InStr(g, h) Then
'75% terinfeksi virus sality
    If FseC = SalityA Then 'pengecekan virtual adrres
          '100 % terinfeksi salityA
            viruzE = "InFected W32.SalityA"
            tipeVirus = "Heuristic"
            CekDNA = True: GoTo r
    ElseIf FseC = SalityB Then
            viruzE = "InFected W32.SalityB"
            tipeVirus = "Heuristic"
            CekDNA = True: GoTo r
    End If
    
End If
End If
r:
End Function
Public Function cekverHDR(fPN As String) As Boolean
Dim a As VERHEADER
Dim i As Long
GetVerHeader fPN, a
For i = 0 To jmlVER
 If InStr(UCase$(a.CompanyName), VirVERid(i)) Then
    viruzE = virVERname(i)
    tipeVirus = "Heuristic"
    cekverHDR = True: GoTo Aaa
 End If
Next i
Aaa:
End Function
Public Function CekFisiK(Bin As String) As Boolean
Dim sKrip As String
CekFisiK = False
sKrip = StrConv(Bin, vbUpperCase)
If InStr(sKrip, StrReverse("TCEJBOETAERC")) > 0 Then

If InStr(sKrip, StrReverse("ETIRWGER")) > 0 Then viruzE = "Malcode.script(regwrite)": GoTo bawah
If InStr(sKrip, StrReverse("RGMKSATELBASID")) > 0 Then viruzE = "Malcode.script(DisableTaskMan)": GoTo bawah
If InStr(sKrip, StrReverse("ESREVERRTS")) > 0 Then viruzE = "Malcode.script(Encripted)": GoTo bawah
If InStr(sKrip, StrReverse("ELIFYPOC")) > 0 Then viruzE = "Malcode.script(CopyFile)": GoTo bawah
If InStr(sKrip, StrReverse("NUROTUA")) > 0 Then viruzE = "Malcode.script(Autorun)": GoTo bawah
If InStr(sKrip, StrReverse("REDLOFLAICEPSTEG")) > 0 Then viruzE = "Malcode.script(SpecialDir)": GoTo bawah
If InStr(sKrip, StrReverse("SURIV")) > 0 Then viruzE = "Malcode.script(VIRUS)": GoTo bawah

End If
sKrip = vbNullString
Exit Function
bawah:
tipeVirus = "Heuristic"
sKrip = vbNullString
CekFisiK = True
End Function
Public Function ceksorCuT(Bin As String) As Boolean
Dim sKrip As String
ceksorCuT = False
sKrip = StrConv(Bin, vbUpperCase)
If InStr(sKrip, "WSCRIPT.EXE") > 0 Then GoTo bawah
If InStr(sKrip, ".DLS") > 0 Then GoTo bawah
If InStr(sKrip, ".VBS") > 0 Then GoTo bawah
If InStr(sKrip, RamnitSrc) > 0 Then GoTo bawah

sKrip = vbNullString
Exit Function
bawah:
viruzE = "Junx.Shortcut"
tipeVirus = "Heuristic"
ceksorCuT = True
sKrip = vbNullString
End Function
Public Function CekIconB(Pasien As String, LbyTe As Long) As Boolean
Dim IconIDNow As String, i As Long
CekIconB = False
IconIDNow = CalcIcon(Pasien, LbyTe)
    If InStrRev(IntViriIconID, IconIDNow) <> 0 Then
    For i = 0 To jmlIco
        If ViriIconID(i) = IconIDNow Then
        viruzE = ViriiconNa(i)
        tipeVirus = "Heuristic"
        CekIconB = True
        End If
    Next i
    End If
IconIDNow = vbNullString
End Function
Public Function ceKAUT(ByVal paT As String, Bin As String) As Boolean
Dim nmfile As String
Dim strDRV As String
Dim isiau As String
nmfile = UCase$(Mid$(paT, 4))
strDRV = VBA.Left$(paT, 3)
If FileAdaX(strDRV & "autorun.inf") = True Then
isiau = StrConv(GFQ(strDRV & "autorun.inf"), vbUpperCase)
 If InStr(isiau, nmfile) > 0 Then
    If VBA.Left$(Bin, 2) = "MZ" Then
      ReadPE StrConv(Bin, vbFromUnicode)
      PEtemp = dWptR
    End If
    ceKAUT = True
    viruzE = "Autorun.Infector"
    tipeVirus = "Heuristic"
 Else
    ceKAUT = False
 End If
End If
End Function
Public Function cekunrealX(Where As String) As Boolean
If InStr(Where, "system32\BlAcKLiSt") > 0 Then
If InStr(Where, "By unrealx33") > 0 Then
            cekunrealX = True
            viruzE = "unrealx33.BAT"
            tipeVirus = "Worm"
End If
End If
End Function
Public Function cekDocinF(Where As String) As Boolean
If InStr(Where, "ÐÏà¡±á") > 0 Then
If InStr(Where, "Word.Document.8") > 0 Then
            cekDocinF = True
            viruzE = "Infected Document"
            tipeVirus = "Virus"
End If
End If
End Function
Public Function cekVirutGen(Where As String) As Boolean
If InStr(Where, "…Àt!Pÿ•$@") > 0 Then
            cekVirutGen = True
            viruzE = "Virut.Gen"
            tipeVirus = "Virus"
End If
End Function
Public Function cekCanTix(Where As String) As Boolean
If InStr(Where, "[autorun]") > 0 Then
If InStr(Where, "open=WScript.exe //e:VBScript dekstop.ini auto") > 0 Then
            cekCanTix = True
            viruzE = "Cantix.Inf"
            tipeVirus = "Worm"
End If
End If
End Function
Public Function cEkEICAR(Where As String) As Boolean
If InStr(Where, "X5O!P%@AP[4\PZX54(P^)7CC)7}$") > 0 Then
If InStr(Where, "EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*") > 0 Then
            cEkEICAR = True
            viruzE = "EICAR.Test"
            tipeVirus = "Virus Test"
End If
End If
End Function
Public Function cekRunouce(Where As String) As Boolean
If InStr(Where, "MAIL FROM: imissyou@btamail.net.cn") > 0 Then
            cekRunouce = True
            viruzE = "Runouce.B2"
            tipeVirus = "Virus"
End If
End Function
Public Function cekPirut(Where As String) As Boolean
If InStr(Where, "‹Oý¿") > 0 Then
            cekPirut = True
            viruzE = "Win32.Sality.AT"
            tipeVirus = "Heurictic"
End If
End Function
Public Function cekTrojanGen(dataQ As String) As Boolean
If InStr(dataQ, REGrun) > 0 Then
If InStr(dataQ, REGhiden) > 0 Then
            cekTrojanGen = True
            viruzE = "Trojan.Generic"
            tipeVirus = "Heuristic"
End If
End If
End Function
Public Function cekAlman(Where As String) As Boolean
If InStr(Where, "J6vÙ0fc›") - InStr(Where, "16äØøcc‘") = 9 Then
        cekAlman = True
        viruzE = "Infected Alman.A"
        tipeVirus = "Virus"
End If
End Function
Public Function lenFileEX(Pathx As String) As Long
Dim heFile As Long, zizE As Long
 heFile = CreateFileW(StrPtr(Pathx), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
 If heFile > 0 Then
    zizE = GetFileSize(heFile, 0)
    CloseHandle heFile
    lenFileEX = zizE
 End If
heFile = vbNull
zizE = vbNull
End Function
Public Function FileAdaX(ByVal FilNAM As String) As Boolean
Dim heFile As Long
FileAdaX = False
 heFile = CreateFileW(StrPtr(FilNAM), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
 If heFile > 0 Then
 FileAdaX = True
 CloseHandle heFile
 End If
End Function
Public Sub InfoDB(jVPE As Long, nVPE() As String, jVIC As Long, nVIC() As String)
jVPE = jmlVPE
nVPE() = virPEname()
jVIC = jmlIco
nVIC() = ViriiconNa()
End Sub
Public Function GFQ(strFilePath As String, Optional bolAsString = True)
  Dim arrFileMain() As Byte
  Dim lngSize As Long, lngRet As Long
  Dim lngFileHandle As Long
  Dim ofData As OFSTRUCT, OF_READ As Long
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
