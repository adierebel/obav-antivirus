VERSION 5.00
Begin VB.Form fRtpSystem 
   ClientHeight    =   2235
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   3090
   Icon            =   "fRtpSystem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Tstartup 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Tguard 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tRtp 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label DRibeLBL 
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Begin VB.Menu openscan 
         Caption         =   "Open ObAV Scanner"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu Pausegrd 
         Caption         =   "Pause Guard"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit Guard"
      End
      Begin VB.Menu sguydhrtg 
         Caption         =   "-"
      End
      Begin VB.Menu asas 
         Caption         =   "Check for Update"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "fRtpSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L1(100) As String
Dim L2(100) As String
Dim tL1 As Long, tL2 As Long

Dim iewindow As InternetExplorer
Dim currentwinDows  As New ShellWindows
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Dim rtpAKTIV As Boolean
Dim guardAKTIV As Boolean
Private Sub About_Click()
Fabout.Show
End Sub
Private Sub asas_Click()
OpenURL "http://www.obav.net", Me.hwnd
End Sub
Private Sub exit_Click()
If MsgBox("Are you sure to exit for Protections ObAV Guard ?", vbYesNo + vbQuestion + vbSystemModal) = vbYes Then
ShellEXE fixP(ObAVdir) & "ObAVSvc.exe", "-stop"
bsTray.ShellTrayIconRemove fRtpSystem.hwnd
Call unHook
ExitProcess 0
End If
End Sub
Private Sub Form_Load()
Me.Caption = vbNullChar
App.Title = vbNullChar
End Sub
Sub GuardSTART()
'If FindWindow("#32770", "obav_22091993") > 1 Then
'MsgBox "ObAV Guard is Running", vbCritical + vbSystemModal: End
'End If
quiteSTATE.Poke "ojanBLANK IS THE BEST"
Tguard.Enabled = True
rtpSTATE.Poke "aktiv"
guardSTATE.Poke "aktiv"
rtpAKTIV = True
guardAKTIV = True
tRtp.Enabled = True
End Sub
Private Sub callRTP()
tRtp.Enabled = True
rtpSTATE.Poke "aktiv"
bsTray.ShellTrayBalloonTipShow hwnd, 1, verobAV, "ObAV Explorer Guard Activated"
End Sub
Private Sub stopRTP()
tRtp.Enabled = False
rtpSTATE.Poke "non aktiv"
bsTray.ShellTrayBalloonTipShow hwnd, 1, verobAV, "ObAV Explorer Guard NonActivated"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ShellEXE fixP(ObAVdir) & "ObAVSvc.exe", "-stop"
Call unHook
ExitProcess 0
End Sub
Private Sub Form_Terminate()
ShellEXE fixP(ObAVdir) & "ObAVSvc.exe", "-stop"
Call unHook
ExitProcess 0
End Sub
Private Sub openscan_Click()
ShellEXE appFullpatH
End Sub
Private Sub Pausegrd_Click()
If Pausegrd.Checked Then
rtpSTATE.Poke "aktiv"
guardSTATE.Poke "aktiv"
Pausegrd.Checked = False
Else
rtpSTATE.Poke "non aktiv"
guardSTATE.Poke "non aktiv"
Pausegrd.Checked = True
End If
End Sub
Sub Tguard_Timer()
Dim rtpTMP As Boolean
Dim guardTMP As Boolean
rtpTMP = rtpAKTIV
guardTMP = guardAKTIV
Call cekMEMrtp
Call cekMEMGuard
Call cekMEMquit
If rtpAKTIV Then
If rtpTMP = False Then callRTP
Else
If rtpTMP = True Then stopRTP
End If

If guardAKTIV Then
If guardTMP = False Then pauseHOOK False
Else
If guardTMP = True Then pauseHOOK True
End If
End Sub

Private Sub Timer1_Timer()
Dim sDriveName          As String
Dim DriveLabel          As String
Dim nDriveNameLen       As Long
Dim nilanyA As String
Dim fileN As String
If AdakahFDBaru(LastFlashVolume) = True Then
   nDriveNameLen = 128
   sDriveName = String$(nDriveNameLen, 0)
   If GetVolumeInformationW(StrPtr(Chr(LastFlashVolume) & ":\"), StrPtr(sDriveName), nDriveNameLen, ByVal 0, ByVal 0, ByVal 0, 0, 0) Then
       DriveLabel = Left$(sDriveName, InStr(1, sDriveName, ChrW$(0)) - 1)
   Else
       DriveLabel = vbNullString
   End If
DRibeLBL.Caption = DriveLabel
fileN = fixP(ObAVdir) & "ObAV.cfg"
nilanyA = GetINI("Options", "ScanFD", fileN)
If nilanyA = 1 Then
ProcScanRTP Chr(LastFlashVolume) & ":\", False
'MsgBox ("HAHAH" & " [ " & DriveLabel & " (" & Chr(LastFlashVolume) & ") ] ?"), vbExclamation, ""
End If
End If
End Sub
Private Sub tRtp_Timer()
Dim i As Long
Dim k As Long
Dim currentlocation As String

On Error GoTo TheEnd
If currentwinDows.Count > 0 Then
Erase L2
tL2 = 0
    For Each iewindow In currentwinDows
        DoEvents
        If iewindow.Busy Then GoTo busysignal
        currentlocation = iewindow.LocationURL
        If Mid$(currentlocation, 1, 7) = "file://" Then
                 currentlocation = Replace(currentlocation, "file:///", "")
                 currentlocation = Replace(currentlocation, "%20", " ")
                 currentlocation = Replace(currentlocation, "/", "\")
                 currentlocation = Replace(currentlocation, "%5B", "[")
                 currentlocation = Replace(currentlocation, "%5D", "]")
         For k = 0 To tL2 - 1
            If currentlocation = L2(k) Then GoTo busysignal
         Next k
         L2(tL2) = currentlocation
         tL2 = tL2 + 1
         For k = 0 To tL1 - 1
            If currentlocation = L1(k) Then GoTo busysignal
         Next k
         Cari = True
         scanN currentlocation, True, False, 2
         rtpPATH.Poke currentlocation
End If
busysignal:
    Next
    Erase L1
    tL1 = 0
    For k = 0 To tL2 - 1
        L1(k) = L2(k)
        tL1 = tL1 + 1
    Next k
    End If
TheEnd:
End Sub
Private Sub cekMEMrtp()
If Trim$(TrimW(rtpSTATE.Peek)) = "aktiv" Then
rtpAKTIV = True
Else
rtpAKTIV = False
End If
End Sub
Private Sub cekMEMGuard()
If Trim$(TrimW(guardSTATE.Peek)) = "aktiv" Then
guardAKTIV = True
Else
guardAKTIV = False
End If
End Sub
Private Sub cekMEMquit()
If Trim$(TrimW(quiteSTATE.Peek)) = "1234567890" Then
ShellEXE fixP(ObAVdir) & "ObAVSvc.exe", "-stop"
bsTray.ShellTrayIconRemove fRtpSystem.hwnd
Call unHook
ExitProcess 0
End If
End Sub
