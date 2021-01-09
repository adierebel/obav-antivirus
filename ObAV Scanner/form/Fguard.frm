VERSION 5.00
Begin VB.Form Fguard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   Icon            =   "Fguard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Block+Clean"
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Block"
      Height          =   495
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Allow"
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   5895
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      Picture         =   "Fguard.frx":000C
      ScaleHeight     =   3255
      ScaleWidth      =   6855
      TabIndex        =   4
      Top             =   0
      Width           =   6855
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
Private Sub Form_Load()
Dim i As Integer
Dim nilanyA As String
Dim fileN As String
Me.Caption = vbNullChar
FormCenter Me
onTOP hwnd
fileN = fixP(App.path) & "ObAV.cfg"
nilanyA = GetINI("Options", "Sound", fileN)
If nilanyA = 1 Then
SoundBuffer = LoadResData("THREAT", "SOUND")
sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End If
End Sub
Sub showREPORT(nama As String, path As String, tipe As String, pId As Long)
Me.Show
loADingF = True
pidhOOK = pId
pathHook = path
Text1.Text = "ObAV Mendeteksi Adanya Malware/virus yang akan anda jalankan" & vbCrLf & _
"Threat Name : " & nama & vbCrLf & "Threat Path   : " & path & vbCrLf & "Detection      : " & tipe
BEEEP
End Sub

Private Sub Form_Unload(Cancel As Integer)
FgrdEnd = True
loADingF = False
End Sub
