VERSION 5.00
Begin VB.Form Fwizard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6675
   Icon            =   "Fwizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Fwizard.frx":591A
   ScaleHeight     =   4590
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Index           =   3
      Left            =   1920
      ScaleHeight     =   3435
      ScaleWidth      =   4515
      TabIndex        =   8
      Top             =   480
      Width           =   4575
      Begin VB.PictureBox progxxx 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   4035
         TabIndex        =   12
         Top             =   2160
         Width           =   4095
         Begin VB.PictureBox progress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   345
            ScaleWidth      =   15
            TabIndex        =   13
            Top             =   0
            Width           =   40
         End
      End
      Begin VB.Label Lst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Index           =   2
      Left            =   1920
      ScaleHeight     =   3435
      ScaleWidth      =   4515
      TabIndex        =   6
      Top             =   480
      Width           =   4575
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "Fwizard.frx":11947
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Index           =   1
      Left            =   1920
      ScaleHeight     =   3435
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   480
      Width           =   4575
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "Fwizard.frx":1194D
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to ObAV Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   4215
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
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Dim resOurce() As Byte
Select Case Index
    Case 0
        pDisab
        tom = tom - 1
        Picture1(tom).Visible = True
        If tom = 1 Then _
            Command1(0).Enabled = False
        If tom < 3 Then _
            Command1(1).Enabled = True
    Case 1
            If isADMIN = False Then
                MsgBox "Please Run This Program As ADMINISTRATOR", vbCritical + vbSystemModal
                Lst.Caption = "Error!,No Administrator Priviliges"
                Picture1(3).Visible = True
                Command1(1).Enabled = False
                GoTo exeiTs
            End If
            pDisab
            tom = tom + 1
            Picture1(tom).Visible = True
            If tom > 1 Then _
                Command1(0).Enabled = True
            If tom > 2 Then
                Command1(2).Enabled = False
                Command1(1).Enabled = False
                Command1(0).Enabled = False
                Lst.Caption = "Setup is Loading Files"
                tunggON 20
                progbar progress, 10
                '---------------------------
                Lst.Caption = "Installing File : Uninstall.exe"
                If FolderADA(ObAVdir) = False Then
                    MkDir ObAVdir
                End If
                CopyFileW StrPtr(appFullpatH), StrPtr(fixP(ObAVdir) & "Uninstall.exe"), 0
                tunggON 15
                Lst.Caption = "Installing File : KprocMon.sys"
                resOurce = LoadResData(1, "A")
                encripFile resOurce, fixP(ObAVdir) & "KprocMon.sys"
                'Open fixP(ObAVdir) & "KprocMon.sys" For Binary Access Write As #1
                '    Put #1, , resOurce
                'Close #1
                tunggON 15
                progbar progress, 10
                Lst.Caption = "Installing File : ObavExt.dll"
                resOurce = LoadResData(2, "A")
                encripFile resOurce, fixP(ObAVdir) & "ObavExt.dll"
                'Open fixP(ObAVdir) & "ObavExt.dll" For Binary Access Write As #1
                '    Put #1, , resOurce
                'Close #1
                tunggON 15
                progbar progress, 10
                Lst.Caption = "Installing File : ObavSpk.sys"
                resOurce = LoadResData(3, "A")
                encripFile resOurce, fixP(ObAVdir) & "ObavSpk.sys"
                'Open fixP(ObAVdir) & "ObavSpk.sys" For Binary Access Write As #1
                '    Put #1, , resOurce
                'Close #1
                tunggON 15
                progbar progress, 10
                Lst.Caption = "Installing File : ObavScanner.exe"
                resOurce = LoadResData(4, "A")
                encripFile resOurce, fixP(ObAVdir) & "ObavScanner.exe"
                'Open fixP(ObAVdir) & "ObAV.exe" For Binary Access Write As #1
                '    Put #1, , resOurce
                'Close #1
                tunggON 15
                progbar progress, 10
                
                Lst.Caption = "Installing File : ObavLoader.exe"
                resOurce = LoadResData(5, "A")
                encripFile resOurce, fixP(ObAVdir) & "ObavLoader.exe"
                'Open fixP(ObAVdir) & "ObAVLoader.exe" For Binary Access Write As #1
                '    Put #1, , resOurce
                'Close #1
                tunggON 15
                progbar progress, 10
            
                Lst.Caption = "Installing File : ObavSvc.exe"
                resOurce = LoadResData(6, "A")
                encripFile resOurce, fixP(ObAVdir) & "ObavSvc.exe"
                'Open fixP(ObAVdir) & "ObAVSvc.exe" For Binary Access Write As #1
                '    Put #1, , resOurce
                'Close #1
                tunggON 15
                progbar progress, 10
                
                Lst.Caption = "Installing File : ObavGuard.exe"
                resOurce = LoadResData(7, "A")
                encripFile resOurce, fixP(ObAVdir) & "ObavGuard.exe"
                'Open fixP(ObAVdir) & "ObAVLoader.exe" For Binary Access Write As #1
                '    Put #1, , resOurce
                'Close #1
                tunggON 15
                progbar progress, 10
                
                Lst.Caption = "Installing Runtime"
                CopyFileW StrPtr(fixP(App.path) + "msvbvm60.dll"), StrPtr(fixP(ObAVdir) & "msvbvm60.dll"), 0
                tunggON 15
                progbar progress, 5
                '---------------------------
                Lst.Caption = "Registering File"
                ShellEXE fixP(ObAVdir) & "ObavSvc.exe", "-install"
                tunggON 15
                progbar progress, 10
                '---------------------------
                Lst.Caption = "Updating your System"
                'SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", ChrW$(&H2126) & "bAV Guard", fixP(ObAVdir) & "ObAV.exe"
                SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "DisplayName", "ObAV"
                SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "DisplayIcon", fixP(ObAVdir) & "ObAV.exe"
                SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "Publisher", "ObAV Team."
                SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "HelpLink", "http://facebook.com/obavantivirus"
                SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "UninstallString", fixP(ObAVdir) & "Uninstall.exe"
                RegObavExt True
                Call SetIniDefault
                Call obavShortCUT
                ShellEXE fixP(ObAVdir) & "ObavSvc.exe", "-start"
                tunggON 5
                Call BuildServisSPK(True)
                Call BKproc(True)
                tunggON 15
                progbar progress, 5
                '---------------------------'

                'ShellEXE fixP(ObAVdir) & "ObAV.exe", "-rtp"
                Lst.Caption = "Finish"
exeiTs:
                Command1(2).Caption = "Finish"
                Command1(2).Enabled = True
            End If
    Case 2
        'unloD = True
        'Unload Me
        ExitProcess 0
End Select
End Sub
Private Sub Form_Load()
tom = 1
Me.Caption = vbNullChar
pDisab
Picture1(1).Visible = True
Text2.Text = "  Instalations Progress" & vbCrLf & vbCrLf & "  1.Loading File & Component" & _
            vbCrLf & "  2.Installing File" & vbCrLf & "  3.Registering File" & vbCrLf & _
            "  4.Updating your System"
FormCenter Me
End Sub
Private Function pDisab()
Dim i As Long
For i = 1 To Picture1.Count
Picture1(i).Visible = False
Next i
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If unloD = False Then Cancel = 1
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
LiNk.TargetPath = fixP(ObAVdir) & "ObavLoader.exe"
LiNk.Arguments = fixP(ObAVdir) & "ObavScanner.exe"
LiNk.save
Set LiNk = Nothing

starprogDIR = fixP(GetSpecFolder(STAR_PROGRAMS)) & "ObAV AntiVirus"
If FolderADA(starprogDIR) = False Then
MkDir starprogDIR
End If
Set LiNk = wS.cReateShortCUT(fixP(starprogDIR) & "ObAV.lnk")
LiNk.Description = "ObAV Scanner"
LiNk.TargetPath = fixP(ObAVdir) & "ObavLoader.exe"
LiNk.Arguments = fixP(ObAVdir) & "ObavScanner.exe"
LiNk.save
Set LiNk = Nothing

Set LiNk = wS.cReateShortCUT(fixP(starprogDIR) & "Uninstall.lnk")
LiNk.Description = "Uninstall ObAV"
LiNk.TargetPath = fixP(ObAVdir) & "Uninstall.exe"
LiNk.save
Set LiNk = Nothing

Set wS = Nothing
End Sub
