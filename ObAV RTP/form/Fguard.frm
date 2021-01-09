VERSION 5.00
Begin VB.Form Fguard 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   Icon            =   "Fguard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Fguard.frx":000C
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   6705
      TabIndex        =   4
      Top             =   2520
      Width           =   6735
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   375
         Index           =   8
         Left            =   1680
         TabIndex        =   13
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   1800
         TabIndex        =   8
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detection"
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
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Threat Path"
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
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Threat Name"
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
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Block+Clean"
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Block"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Allow"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "....."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ObAV Executable Guard mendeteksi program yang anda jalankan sebagai Malware."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Threat Detected !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Fguard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pidhOOK As String
Private pathHook As String
Sub showREPORT(nama As String, path As String, tipe As String, pId As Long)
Me.Show
loADingF = True
pidhOOK = pId
pathHook = path
Label1(4).Caption = ": " & nama
Text2.Text = path
Label1(5).Caption = ": " & tipe
End Sub
Private Sub Form_Load()
Dim lonRecT As String
lonRecT = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 10, 10)
SetWindowRgn Me.hWnd, lonRecT, True
Dim i As Integer
Dim nilanyA As String
Dim fileN As String
Me.Caption = vbNullChar
FormCenter Me
onTOP hWnd
fileN = fixP(App.path) & "ObAV.cfg"
nilanyA = GetINI("Options", "Sound", fileN)
If nilanyA = 1 Then
SoundBuffer = LoadResData("THREAT", "SOUND")
sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        MoveForm Me.hWnd
    End If
End Sub

Private Sub Timer1_Timer()
If Label1(0).ForeColor = &HFFFFFF Then
Label1(0).ForeColor = &HFF&
Else
Label1(0).ForeColor = &HFFFFFF
End If
End Sub
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
FgrdEnd = True
Case 1
Killproc pidhOOK, True
FgrdEnd = True
Case 2
Killproc pidhOOK, True
Killproc pidhOOK, True
EksekusiFile "karantina", pathHook
FgrdEnd = True
End Select

loADingF = False
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
FgrdEnd = True
loADingF = False
End Sub
Private Sub Timer2_Timer()
Label1(7).Caption = Day(Date) & "/" & Month(Date) & "/" & Year(Date) & " " & " " & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
End Sub
