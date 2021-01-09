VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin ObAV.uTabSonny uTabSonny123 
      Height          =   4905
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8652
      tabcount        =   2
      judul(1)        =   "General Setting"
      judul(2)        =   "Scan Setting"
      Begin VB.PictureBox Pictur22e1 
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4275
         ScaleWidth      =   6675
         TabIndex        =   2
         Top             =   435
         Width           =   6735
         Begin ObAV.ucFrame ucFrame2sd 
            Height          =   1455
            Left            =   240
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2566
            Caption         =   "Window Setting's"
            Begin VB.CheckBox StArTup 
               Caption         =   "Show Startup Screen"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   1080
               Width           =   5895
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Show On top Window"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   11
               Top             =   720
               Width           =   5895
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Anti Destroy Window"
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   10
               Top             =   360
               Width           =   5895
            End
         End
         Begin ObAV.ucFrame ucFrame1fgh 
            Height          =   1455
            Left            =   240
            Top             =   1800
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2566
            Caption         =   "Extra Setting's"
            Begin VB.CheckBox snd 
               Caption         =   "Sound Warning"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   720
               Width           =   5895
            End
            Begin VB.CheckBox ScanFD 
               Caption         =   "Automatic Scan Removable Disk"
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   1080
               Width           =   5895
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Scan with in Explorer"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   13
               Top             =   360
               Width           =   5895
            End
         End
         Begin ObAV.ucFrame ucFrame1 
            Height          =   735
            Left            =   240
            Top             =   3360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   1296
            Caption         =   "Info"
            Begin VB.Label infolbl 
               Caption         =   "General setting for ObAV"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   320
               Width           =   6015
            End
         End
      End
      Begin VB.PictureBox Picture222 
         Height          =   4335
         Left            =   -13615
         ScaleHeight     =   4275
         ScaleWidth      =   6675
         TabIndex        =   1
         Top             =   435
         Width           =   6735
         Begin ObAV.ucFrame ucFrame3ff 
            Height          =   735
            Left            =   240
            Top             =   3360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   1296
            Caption         =   "Info"
            Begin VB.Label LblInfop 
               Caption         =   "Setting for Scanning file process."
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   320
               Width           =   6015
            End
         End
         Begin ObAV.ucFrame ucFrame1ffcxv 
            Height          =   3015
            Index           =   4
            Left            =   240
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5318
            Caption         =   "Scan Setting's"
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Heuristic Sortcut (Yuyun, Serviks, etc)"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   9
               Top             =   2520
               Value           =   1  'Checked
               Width           =   3735
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Heuristic MalScript/VBS"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   8
               Top             =   2160
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Heuristic Version Header (Fake File)"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   7
               Top             =   1800
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Suspect VMX Extention (Like Conficker)"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   6
               Top             =   1440
               Value           =   1  'Checked
               Width           =   4095
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Heuristic Icon"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   5
               Top             =   1080
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Heuristic PE Header (Sality, Virus, etc)"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   4
               Top             =   720
               Value           =   1  'Checked
               Width           =   3735
            End
            Begin VB.CheckBox chkOP 
               Caption         =   "Enable Heuristic String (Alman, Doc Infected)"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Value           =   1  'Checked
               Width           =   3735
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkOP_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
LblInfop.Caption = "Enable or Disable Heuristic String (Alman, Doc Infected)"
Case 1
LblInfop.Caption = "Enable or Disable Heuristic PE Header (Sality, Virus, etc)"
Case 3
LblInfop.Caption = "Enable or Disable Heuristic Icon"
Case 4
LblInfop.Caption = "Enable or Disable Suspect VMX Extention (Like Conficker)"
Case 5
LblInfop.Caption = "Enable or Disable Heuristic Version Header (Fake File)"
Case 6
LblInfop.Caption = "Enable or Disable Heuristic MalScript/VBS"
Case 7
LblInfop.Caption = "Enable or Disable Heuristic Sortcut (Yuyun, Serviks, etc)"
'====
Case 9
infolbl.Caption = "Enable or Disable Anti Destroy Window Mode"
Case 2
infolbl.Caption = "Enable or Disable Show On top Window Mode"
Case 8
infolbl.Caption = "Enable or Disable Shell Scan with in Explorer"
End Select
End Sub
Private Sub Pictur22e1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "General setting for ObAV"
End Sub
Private Sub Picture222_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblInfop.Caption = "Setting for Scanning file process."
End Sub
Private Sub ScanFD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "Enable or Disable Automatic Scan Removable Disk"
End Sub
Private Sub snd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "Enable or Disable Sound Warning"
End Sub
Private Sub StArTup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "Enable or Disable Show Startup Screen"
End Sub

