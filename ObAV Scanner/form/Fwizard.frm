VERSION 5.00
Begin VB.Form Fwizard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "Fwizard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Fwizard.frx":000C
   ScaleHeight     =   4320
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   240
      Top             =   360
   End
   Begin VB.PictureBox welcome 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   5775
      TabIndex        =   4
      Top             =   1200
      Width           =   5775
      Begin VB.PictureBox installhover 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1080
         MouseIcon       =   "Fwizard.frx":A7CE
         MousePointer    =   99  'Custom
         Picture         =   "Fwizard.frx":A920
         ScaleHeight     =   735
         ScaleWidth      =   3420
         TabIndex        =   8
         Top             =   1080
         Width           =   3420
      End
      Begin VB.PictureBox Scannerhover 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1080
         MouseIcon       =   "Fwizard.frx":FECA
         MousePointer    =   99  'Custom
         Picture         =   "Fwizard.frx":1001C
         ScaleHeight     =   735
         ScaleWidth      =   3420
         TabIndex        =   7
         Top             =   1920
         Width           =   3420
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "silahkan pilih opsi di bahwah ini."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -120
         TabIndex        =   6
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Komputer Anda belum terinstall ObAV Guard, "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -120
         TabIndex        =   5
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Index           =   3
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   1200
      Width           =   5775
      Begin ObAV.ucProgressBar ucPB 
         Height          =   375
         Left            =   120
         Top             =   2160
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Lst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   5535
      End
   End
   Begin VB.PictureBox Thanks 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   5775
      TabIndex        =   12
      Top             =   1200
      Width           =   5775
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   2040
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1830
         Left            =   960
         Picture         =   "Fwizard.frx":15B4D
         ScaleHeight     =   1830
         ScaleWidth      =   4065
         TabIndex        =   13
         Top             =   360
         Width           =   4065
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Otomatis keluar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1815
      End
   End
   Begin VB.PictureBox InstallBro 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   5775
      TabIndex        =   9
      Top             =   1200
      Width           =   5775
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "Fwizard.frx":1CFD8
         Top             =   120
         Width           =   5535
      End
      Begin VB.PictureBox PasangHover 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1200
         MouseIcon       =   "Fwizard.frx":1D89C
         MousePointer    =   99  'Custom
         Picture         =   "Fwizard.frx":1D9EE
         ScaleHeight     =   735
         ScaleWidth      =   3420
         TabIndex        =   11
         Top             =   2040
         Width           =   3420
      End
   End
End
Attribute VB_Name = "Fwizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private winHwnd As Long
Dim JmlPersen As Integer
Private tom As Long
Private unloD As Boolean
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
Private Sub pangg()
aKsiPepsodent (3)
On Error Resume Next
Dim resOurce() As Byte
        Lst.Caption = "Setup is Loading Files"
        tunggON 20
        ucPB.Value = ucPB.Value + 10
        '---------------------------
        Lst.Caption = "Installing File : ObAV.exe"
        If FolderADA(ObAVdir) = False Then
        MkDir ObAVdir
        End If
        CopyFileW StrPtr(appFullpatH), StrPtr(fixP(ObAVdir) & "ObAV.exe"), 0
        tunggON 15
        Lst.Caption = "Installing File : KprocMon.sys"
        resOurce = LoadResData(1, "A")
        Open fixP(ObAVdir) & "KprocMon.sys" For Binary Access Write As #1
        Put #1, , resOurce
        Close #1
        tunggON 15
        ucPB.Value = ucPB.Value + 10
        Lst.Caption = "Installing File : ObavExt.dll"
        resOurce = LoadResData(2, "A")
        Open fixP(ObAVdir) & "ObavExt.dll" For Binary Access Write As #1
        Put #1, , resOurce
        Close #1
        tunggON 15
        ucPB.Value = ucPB.Value + 10
        Lst.Caption = "Installing File : ObavSpk.sys"
        resOurce = LoadResData(3, "A")
        Open fixP(ObAVdir) & "ObavSpk.sys" For Binary Access Write As #1
        Put #1, , resOurce
        Close #1
        tunggON 15
        ucPB.Value = ucPB.Value + 10
        Lst.Caption = "Installing File : ObavKbp.sys"
        resOurce = LoadResData(4, "A")
        Open fixP(ObAVdir) & "ObavKbp.sys" For Binary Access Write As #1
        Put #1, , resOurce
        Close #1
        tunggON 15
        ucPB.Value = ucPB.Value + 10
        '---------------------------
        Lst.Caption = "Registering File"
        tunggON 15
        ucPB.Value = ucPB.Value + 25
       '---------------------------
        Lst.Caption = "Updating your System"
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", ChrW$(&H2126) & "bAV Guard", fixP(ObAVdir) & "ObAV.exe"
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "DisplayName", "ObAV"
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "DisplayIcon", fixP(ObAVdir) & "ObAV.exe"
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "Publisher", "ObAV Team."
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "HelpLink", "http://www.obav.net"
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "UninstallString", fixP(ObAVdir) & "ObAV.exe -uninstall"
        RegObavExt True
        Call SetIniDefault
        Call obavShortCUT
        Call BuildServisSPK
        Call BKproc
        Call BKbeeper
        tunggON 15
        ucPB.Value = ucPB.Value + 25
        '---------------------------
        ShellEXE fixP(ObAVdir) & "ObAV.exe", "-rtp"
        Lst.Caption = "Finish"
aKsiPepsodent (4)
        Timer1.Enabled = True
End Sub
Private Sub Form_Load()
Me.Caption = vbNullChar
FormCenter Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If welcome.Visible = True Then
Cancel = 1
metu
ElseIf Picture1(3).Visible = True Then
Cancel = 1
ElseIf InstallBro.Visible = True Then
Cancel = 1
metu
ElseIf Thanks.Visible = True Then
Cancel = 1
metu
End If
End Sub
Private Sub metu()
ExitProcess 0
Unload Me
End Sub
Private Sub SetIniDefault()
Dim fileN As String
fileN = fixP(ObAVdir) & "ObAV.cfg"
SetINI "Options", "chkop0", "1", fileN
SetINI "Options", "chkop1", "1", fileN
SetINI "Options", "chkop2", "0", fileN
SetINI "Options", "chkop3", "1", fileN
SetINI "Options", "chkop4", "1", fileN
SetINI "Options", "chkop5", "1", fileN
SetINI "Options", "chkop6", "1", fileN
SetINI "Options", "chkop7", "1", fileN
SetINI "Options", "chkop8", "1", fileN
SetINI "Options", "chkop9", "0", fileN
SetINI "Options", "Sound", "1", fileN
SetINI "Options", "Splash", "1", fileN
SetINI "Options", "ScanFD", "1", fileN
End Sub
Private Sub obavShortCUT()
Dim wS As Object, LiNk As Object
Dim starprogDIR As String
Set wS = CreateObject(StrReverse("llehS.tpircsW"))
Set LiNk = wS.cReateShortCUT(fixP(GetSpecFolder(DEKSTOP_PATH)) & "ObAV.lnk")
LiNk.Description = "ObAV Scanner"
LiNk.TargetPath = fixP(ObAVdir) & "ObAV.exe"
LiNk.save
Set LiNk = Nothing

starprogDIR = fixP(GetSpecFolder(STAR_PROGRAMS)) & "ObAV AntiVirus"
If FolderADA(starprogDIR) = False Then
MkDir starprogDIR
End If
Set LiNk = wS.cReateShortCUT(fixP(starprogDIR) & "ObAV.lnk")
LiNk.Description = "ObAV Scanner"
LiNk.TargetPath = fixP(ObAVdir) & "ObAV.exe"
LiNk.save
Set LiNk = Nothing

Set LiNk = wS.cReateShortCUT(fixP(starprogDIR) & "Uninstall.lnk")
LiNk.Description = "Uninstall ObAV"
LiNk.TargetPath = "rundll32.exe"
LiNk.Arguments = "shell32.dll,Control_RunDLL appwiz.cpl,Add or Remove Programs"
LiNk.save
Set LiNk = Nothing

Set wS = Nothing
End Sub
Private Sub installhover_Click()
If isADMIN = False Then
    MsgBox "Please Run This Program As ADMINISTRATOR", vbCritical + vbSystemModal
Else
aKsiPepsodent (1)
End If
End Sub
Private Sub PasangHover_Click()
pangg
End Sub
Private Sub Scannerhover_Click()
Fscann.Show
Me.Hide
Timer2.Enabled = True
End Sub
'************************************************************************************'
Private Sub aKsiPepsodent(KradaK As Integer)
Select Case KradaK
Case 0 'posisi default
Scannerhover.Left = 1080
installhover.Left = 1080
Case 1
welcome.Visible = False
Picture1(3).Visible = False
Thanks.Visible = False
InstallBro.Visible = True
Case 2
PasangHover.Left = 1200
Case 3
welcome.Visible = False
Picture1(3).Visible = True
Thanks.Visible = False
InstallBro.Visible = False
Case 4
welcome.Visible = False
Picture1(3).Visible = False
Thanks.Visible = True
InstallBro.Visible = False
End Select
End Sub
Private Sub installhover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
installhover.Left = 1200
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aKsiPepsodent (0)
aKsiPepsodent (2)
End Sub
Private Sub InstallBro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aKsiPepsodent (2)
End Sub
Private Sub PasangHover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PasangHover.Left = 1320
End Sub
Private Sub Scannerhover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Scannerhover.Left = 1200
End Sub
Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aKsiPepsodent (2)
End Sub
Private Sub Timer1_Timer()
If Label3.Caption = 0 Then
metu
Else
Label3.Caption = Label3.Caption - 1
End If
End Sub
Private Sub Timer2_Timer()
Unload Me
End Sub
Private Sub welcome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aKsiPepsodent (0)
End Sub
