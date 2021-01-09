VERSION 5.00
Begin VB.Form FSys 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   630
   ClientWidth     =   2460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FSys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2460
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tRtp 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Tguard 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Tstartup 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   720
      Top             =   0
   End
   Begin VB.CheckBox SysIcon 
      Caption         =   "HAHAHAH"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.PictureBox Flash2 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   480
      Picture         =   "FSys.frx":4C1A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox Flash1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   120
      Picture         =   "FSys.frx":51A4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   300
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Begin VB.Menu openscan 
         Caption         =   "Open ObAV Scanner"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu Pausegrd 
         Caption         =   "ObAV Guard Enable"
         Checked         =   -1  'True
      End
      Begin VB.Menu exit 
         Caption         =   "Exit Guard"
      End
      Begin VB.Menu dfdf 
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
Attribute VB_Name = "FSys"
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

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)

Private nID As NOTIFYICONDATA
Private LastWindowState As Integer

Public Property Let ToolTip(Value As String)
   nID.szTip = Value & vbNullChar
End Property

Public Property Get ToolTip() As String
   ToolTip = nID.szTip
End Property

Public Property Let Interval(Value As Integer)
   UpdateIcon NIM_MODIFY
End Property

Public Property Get Interval() As Integer

End Property
Public Property Let TrayIcon(Value)
   On Error Resume Next
   ' Value can be a picturebox, image, form or string
   Select Case TypeName(Value)
      Case "PictureBox", "Image"
         Me.Icon = Value.Picture
         RaiseEvent TIcon(Me)
      Case "String"
         If (UCase(Value) = "DEFAULT") Then
            Me.Icon = Flash1.Picture
            RaiseEvent TIcon(Me)
         Else
            ' Sting is filename; load icon from picture file.
            Me.Icon = LoadPicture(Value)
            RaiseEvent TIcon(Me)
         End If
      Case Else
         ' It's a form ?
         Me.Icon = Value.Icon
         RaiseEvent TIcon(Me)
   End Select
   If erR.Number <> 0 Then
   UpdateIcon NIM_MODIFY
   End If
End Property

Private Sub SysIcon_Click()
Select Case SysIcon.Value
Case 0 ' disable
Me.Icon = Flash2
   RaiseEvent TIcon(Me)
   UpdateIcon NIM_MODIFY
Case 1 ' enable
Me.Icon = Flash1
   RaiseEvent TIcon(Me)
   UpdateIcon NIM_MODIFY
End Select
End Sub
Private Sub Form_Load()
   Me.Icon = Flash1
   RaiseEvent TIcon(Me)
   Me.Visible = False
   ToolTip = "ObAV Guard ver.1.1"
   UpdateIcon NIM_ADD
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Result As Long
   Dim Msg As Long
   
   ' The Form_MouseMove is intercepted to give systray mouse events.
   If Me.ScaleMode = vbPixels Then
      Msg = X
   Else
      Msg = X / Screen.TwipsPerPixelX
   End If
   
   Select Case Msg
      Case WM_RBUTTONDBLCLK
         RaiseEvent Click("RBUTTONDBLCLK")
      Case WM_RBUTTONDOWN
         RaiseEvent Click("RBUTTONDOWN")
      Case WM_RBUTTONUP
         ' Popup menu: selectively enable items dependent on context.
                
         RaiseEvent Click("RBUTTONUP")
         PopupMenu mnu
      Case WM_LBUTTONDBLCLK
         RaiseEvent Click("LBUTTONDBLCLK")
      Case WM_LBUTTONDOWN
         RaiseEvent Click("LBUTTONDOWN")
      Case WM_LBUTTONUP
         RaiseEvent Click("LBUTTONUP")
      Case WM_MBUTTONDBLCLK
         RaiseEvent Click("MBUTTONDBLCLK")
      Case WM_MBUTTONDOWN
         RaiseEvent Click("MBUTTONDOWN")
      Case WM_MBUTTONUP
         RaiseEvent Click("MBUTTONUP")
      Case WM_MOUSEMOVE
         RaiseEvent Click("MOUSEMOVE")
      Case Else
         RaiseEvent Click("OTHER....: " & Format$(Msg))
   End Select
End Sub

Private Sub FSys_Unload(Cancel As Integer)
   ' Important: remove icon from tray, and unload this form when
   ' the main form is unloaded.
   UpdateIcon NIM_DELETE
   Unload Me
End Sub
Private Sub UpdateIcon(Value As Long)
   ' Used to add, modify and delete icon.
   With nID
      .cbSize = Len(nID)
      .hwnd = Me.hwnd
      .uID = vbNull
      .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
   End With
   Shell_NotifyIcon Value, nID
End Sub
Private Sub About_Click()
Fabout.Show
End Sub
Private Sub asas_Click()
OpenURL "http://www.obav-av.tk", Me.hwnd
End Sub
Private Sub exit_Click()
If MsgBox("Are you sure to exit for Protections ObAV Guard ?", vbYesNo + vbQuestion + vbSystemModal) = vbYes Then
 UpdateIcon NIM_DELETE
Call unHook
ExitProcess 0
End If
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
Tstartup.Enabled = True
End Sub
Private Sub callRTP()
tRtp.Enabled = True
rtpSTATE.Poke "aktiv"
SysIcon.Value = 1
End Sub
Private Sub stopRTP()
tRtp.Enabled = False
rtpSTATE.Poke "non aktiv"
SysIcon.Value = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call unHook
End Sub
Private Sub openscan_Click()
ShellEXE appFullpatH
End Sub
Private Sub Pausegrd_Click()
If Pausegrd.Checked Then
rtpSTATE.Poke "non aktiv"
guardSTATE.Poke "non aktiv"
Pausegrd.Checked = True
Else
Pausegrd.Checked = False
rtpSTATE.Poke "aktiv"
guardSTATE.Poke "aktiv"
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
 UpdateIcon NIM_DELETE
Call unHook
ExitProcess 0
End If
End Sub
Private Sub Tstartup_Timer()
If GetStringW(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", ChrW$(&H2126) & "bAV Guard") <> fixP(ObAVdir) & "ObAV.exe -rtp" Then
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", ChrW$(&H2126) & "bAV Guard", fixP(ObAVdir) & "ObAV.exe -rtp"
End If
End Sub
