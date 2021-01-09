VERSION 5.00
Begin VB.Form FrtpUser 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   Icon            =   "FrtpUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrtpUser.frx":000C
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command 
      Caption         =   "Ignore"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "Clean Checked"
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "Check All"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4500
      Width           =   1215
   End
   Begin ObAVGuard.ucListView ucListView1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      StyleEx         =   37
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
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ObAV Explorer Guard mendeteksi adanya Malware/Virus pada direktori yang sedang Anda buka."
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
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6615
   End
End
Attribute VB_Name = "FrtpUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click()
CekALL Check1, ucListView1
End Sub
Private Sub Command_Click(Index As Integer)
If Index = 0 Then
Eksekusi "karantina", ucListView1
End If
Unload Me
End Sub
Private Sub Form_Load()
Dim lonRecT As String
lonRecT = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 10, 10)
SetWindowRgn Me.hWnd, lonRecT, True
Dim i As Integer
Dim nilanyA As String
Dim fileN As String
Dim Col As cColumns
Me.Caption = vbNullChar
App.Title = vbNullChar
Set Col = ucListView1.Columns
Col.Add , "Threat Name", , , , 170
Col.Add , "Threat Path", , , , 450
Col.Add , "Type", , , , 100
Col.Add , "Info", , , , 250
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
