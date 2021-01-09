VERSION 5.00
Begin VB.Form FrmScanner 
   BackColor       =   &H00E8382F&
   BorderStyle     =   0  'None
   Caption         =   "ObAV Scanner"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13095
   Icon            =   "FrmScanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmScanner.frx":000C
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   300
      Left            =   720
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   54
      Top             =   960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox TabPic 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   8
      Left            =   4410
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   11
      Top             =   1395
      Width           =   1455
      Begin VB.PictureBox TabPic 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   90
         Index           =   9
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   1455
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hidden (0)"
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
         Index           =   4
         Left            =   0
         TabIndex        =   13
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.PictureBox TabPic 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   2955
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   8
      Top             =   1395
      Width           =   1455
      Begin VB.PictureBox TabPic 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   90
         Index           =   5
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   1455
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Registry (0)"
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
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.PictureBox TabPic 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   1500
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   5
      Top             =   1395
      Width           =   1455
      Begin VB.PictureBox TabPic 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   90
         Index           =   2
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   1455
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Virus (0)"
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
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.PictureBox TabPic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   60
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   1395
      Width           =   1455
      Begin VB.PictureBox TabPic 
         BackColor       =   &H00FF2B3A&
         BorderStyle     =   0  'None
         Height          =   90
         Index           =   4
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   1455
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Report"
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
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ScannNow"
      Height          =   1335
      Left            =   7560
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
   End
   Begin ObAVGuard.DirTree DirTree1 
      Height          =   3975
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   1720
      _ExtentY        =   2143
   End
   Begin VB.PictureBox PicturesTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   0
      Left            =   120
      ScaleHeight     =   4815
      ScaleWidth      =   6735
      TabIndex        =   14
      Top             =   1920
      Width           =   6735
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   480
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   120
      End
      Begin VB.PictureBox PicScan 
         BorderStyle     =   0  'None
         Height          =   1845
         Index           =   2
         Left            =   120
         Picture         =   "FrmScanner.frx":E1AE
         ScaleHeight     =   1845
         ScaleWidth      =   1650
         TabIndex        =   35
         Top             =   120
         Width           =   1650
      End
      Begin VB.TextBox path 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "FrmScanner.frx":16422
         Top             =   2760
         Width           =   5295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Stop"
         Height          =   495
         Left            =   4680
         TabIndex        =   30
         Top             =   4200
         Width           =   1815
      End
      Begin ObAVGuard.ucProgressBar progBSC 
         Height          =   375
         Left            =   120
         Top             =   2160
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
      End
      Begin VB.PictureBox PicScan 
         BorderStyle     =   0  'None
         Height          =   1845
         Index           =   1
         Left            =   120
         Picture         =   "FrmScanner.frx":1642F
         ScaleHeight     =   1845
         ScaleWidth      =   1650
         TabIndex        =   34
         Top             =   120
         Width           =   1650
      End
      Begin VB.PictureBox PicScan 
         BorderStyle     =   0  'None
         Height          =   1845
         Index           =   0
         Left            =   120
         Picture         =   "FrmScanner.frx":1E679
         ScaleHeight     =   1845
         ScaleWidth      =   1650
         TabIndex        =   33
         Top             =   120
         Width           =   1650
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   ": More than 2300 with heuristic"
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
         Index           =   7
         Left            =   3480
         TabIndex        =   53
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   ": 11 Jan 1996"
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
         Index           =   6
         Left            =   3480
         TabIndex        =   52
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   ": 22 Nov 2013"
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
         Index           =   5
         Left            =   3480
         TabIndex        =   51
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   ": Ver 2.0.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Detection"
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
         Left            =   1920
         TabIndex        =   49
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Update"
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
         Index           =   2
         Left            =   1920
         TabIndex        =   48
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Release Date"
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
         Index           =   1
         Left            =   1920
         TabIndex        =   47
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   46
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   7
         Left            =   1680
         TabIndex        =   45
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   6
         Left            =   1680
         TabIndex        =   44
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   5
         Left            =   1680
         TabIndex        =   43
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   42
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Total   "
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
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   40
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Object Detected  "
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
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   38
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Scanned "
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
         TabIndex        =   37
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning : "
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
         TabIndex        =   36
         Top             =   2750
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "File :"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.PictureBox PicturesTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   3
      Left            =   120
      ScaleHeight     =   4815
      ScaleWidth      =   6735
      TabIndex        =   19
      Top             =   1920
      Width           =   6735
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select All"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Unhide Check"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         MouseIcon       =   "FrmScanner.frx":26B77
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   4440
         Width           =   1335
      End
      Begin ObAVGuard.ucListView ucListView1 
         Height          =   3735
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6588
         StyleEx         =   1
      End
   End
   Begin VB.PictureBox PicturesTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   2
      Left            =   120
      ScaleHeight     =   4815
      ScaleWidth      =   6735
      TabIndex        =   18
      Top             =   1920
      Width           =   6735
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select All"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Repair Check"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         MouseIcon       =   "FrmScanner.frx":26CC9
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   4440
         Width           =   1335
      End
      Begin ObAVGuard.ucListView ucListView1 
         Height          =   3735
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6588
         StyleEx         =   1
      End
   End
   Begin VB.PictureBox PicturesTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   1
      Left            =   120
      ScaleHeight     =   4815
      ScaleWidth      =   6735
      TabIndex        =   17
      Top             =   1920
      Width           =   6735
      Begin ObAVGuard.ucListView ucListView1 
         Height          =   3735
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6588
         StyleEx         =   1
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clean Check"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         MouseIcon       =   "FrmScanner.frx":26E1B
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   4440
         Width           =   1335
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "File Scanner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1350
      TabIndex        =   20
      Top             =   900
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "FrmScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command5_Click()
prosSCAN
End Sub

Private Sub Command6_Click()
scAn = False
End Sub

Private Sub Form_Load()
Dim lonRecT As String
lonRecT = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 10, 10)
SetWindowRgn Me.hWnd, lonRecT, True
FormCenter Me
DirTree1.LoadTreeDir True
Me.Caption = vbNullChar
App.Title = vbNullChar
TabPic_Click (0)
gambarTable
End Sub
Private Sub gambarTable()
Dim Col As cColumns
With Me
Set Col = .ucListView1(0).Columns
Col.Add , "Threat Name", , , , 1700
Col.Add , "Threat Path", , , , 4500
Col.Add , "Type", , , , 1000
Col.Add , "Info", , , , 2500

Set Col = .ucListView1(1).Columns
Col.Add , "Value Name", , , , 1700
Col.Add , "Key Path", , , , 5000
Col.Add , "Type", , , , 2000

Set Col = .ucListView1(2).Columns
Col.Add , "File Name", , , , 2000
Col.Add , "File Path", , , , 5000

End With
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        MoveForm Me.hWnd
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
resetTabPic (1)
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
Case 1
Me.WindowState = 1
Case 0
Unload Me
End Select
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        MoveForm Me.hWnd
    End If
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub TabPic_Click(Index As Integer)
Select Case Index
Case 0 'report
resetTabPic (0)
TabPic(0).BackColor = vbWhite
TabPic(4).BackColor = &HFF2B3A
PicturesTabs(1).Visible = False
PicturesTabs(2).Visible = False
PicturesTabs(3).Visible = False
PicturesTabs(0).Visible = True


Case 1 'virus
resetTabPic (0)
TabPic(1).BackColor = vbWhite
TabPic(2).BackColor = &HFF2B3A
PicturesTabs(0).Visible = False
PicturesTabs(2).Visible = False
PicturesTabs(3).Visible = False
PicturesTabs(1).Visible = True


Case 3 'registry
resetTabPic (0)
TabPic(3).BackColor = vbWhite
TabPic(5).BackColor = &HFF2B3A
PicturesTabs(1).Visible = False
PicturesTabs(0).Visible = False
PicturesTabs(3).Visible = False
PicturesTabs(2).Visible = True

Case 8 'hidden
resetTabPic (0)
TabPic(8).BackColor = vbWhite
TabPic(9).BackColor = &HFF2B3A
PicturesTabs(1).Visible = False
PicturesTabs(2).Visible = False
PicturesTabs(0).Visible = False
PicturesTabs(3).Visible = True

End Select
End Sub
Private Sub resetTabPic(Index As Integer)
Select Case Index
Case 0
TabPic(0).BackColor = &HE0E0E0
TabPic(4).BackColor = &HC0C0C0
TabPic(1).BackColor = &HE0E0E0
TabPic(2).BackColor = &HC0C0C0
TabPic(3).BackColor = &HE0E0E0
TabPic(5).BackColor = &HC0C0C0
TabPic(8).BackColor = &HE0E0E0
TabPic(9).BackColor = &HC0C0C0
Case 1
If TabPic(0).BackColor = &HE0E0E0 Then
    TabPic(4).BackColor = &HC0C0C0
Else
TabPic(4).BackColor = &HFF2B3A
End If
If TabPic(1).BackColor = &HE0E0E0 Then
    TabPic(2).BackColor = &HC0C0C0
    Else
TabPic(2).BackColor = &HFF2B3A
End If
If TabPic(3).BackColor = &HE0E0E0 Then
    TabPic(5).BackColor = &HC0C0C0
    Else
TabPic(5).BackColor = &HFF2B3A
End If
If TabPic(8).BackColor = &HE0E0E0 Then
    TabPic(9).BackColor = &HC0C0C0
    Else
TabPic(9).BackColor = &HFF2B3A
End If
End Select
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
resetTabPic (1)
Select Case Index
Case 0 'report
TabPic(4).BackColor = &HFF8080
Case 1 'virus
TabPic(2).BackColor = &HFF8080
Case 2 'registry
TabPic(5).BackColor = &HFF8080
Case 4 'hidden
TabPic(9).BackColor = &HFF8080
End Select
End Sub
Private Sub TabPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
resetTabPic (1)
Select Case Index
Case 0 'report
TabPic(4).BackColor = &HFF8080
Case 1 'virus
TabPic(2).BackColor = &HFF8080
Case 3 'registry
TabPic(5).BackColor = &HFF8080
Case 8 'hidden
TabPic(9).BackColor = &HFF8080
End Select
End Sub
Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 0
TabPic_Click (0)
Case 1
TabPic_Click (1)
Case 2
TabPic_Click (3)
Case 4
TabPic_Click (8)
End Select
End Sub

Private Sub Timer1_Timer()
If PicScan(0).Visible = True Then
PicScan(0).Visible = False
PicScan(2).Visible = False
PicScan(1).Visible = True
ElseIf PicScan(1).Visible = True Then
PicScan(0).Visible = False
PicScan(2).Visible = True
PicScan(1).Visible = False
Else
PicScan(0).Visible = True
PicScan(2).Visible = False
PicScan(1).Visible = False
End If
End Sub
Private Sub Timer2_Timer()
On Error GoTo er
Dim i As Long
Label2(2).Caption = scannedF
Label2(3).Caption = jmlVIR
path.Text = LokasiD
i = (100 / Perc) * scannedF
If i <= 100 Then
progBSC.Value = i
End If
er:
End Sub
