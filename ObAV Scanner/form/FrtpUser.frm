VERSION 5.00
Begin VB.Form FrtpUser 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   Icon            =   "FrtpUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   6720
      TabIndex        =   0
      Top             =   900
      Width           =   6775
      Begin VB.CheckBox Check1 
         Caption         =   "Check All"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2100
         Width           =   1215
      End
      Begin VB.CommandButton Command 
         Caption         =   "Clean Checked"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command 
         Caption         =   "Ignore"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   2
         Top             =   2040
         Width           =   1335
      End
      Begin ObAV.ucListView ucListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4471
         StyleEx         =   37
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      Picture         =   "FrtpUser.frx":000C
      ScaleHeight     =   3735
      ScaleWidth      =   7575
      TabIndex        =   5
      Top             =   0
      Width           =   7575
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
Dim i As Integer
Dim nilanyA As String
Dim fileN As String
Dim Col As cColumns
Me.Caption = vbNullChar
App.Title = vbNullChar
Set Col = ucListView1.Columns
Col.Add , "Threat Name", , , , 1700
Col.Add , "Threat Path", , , , 4500
Col.Add , "Type", , , , 1000
Col.Add , "Info", , , , 2500
FormCenter Me
onTOP hwnd
fileN = fixP(App.path) & "ObAV.cfg"
nilanyA = GetINI("Options", "Sound", fileN)
If nilanyA = 1 Then
SoundBuffer = LoadResData("THREAT", "SOUND")
sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End If
End Sub
