VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Full System Area"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "System Area"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quick Scan"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Full Scan"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin ObAVScanner.DirTree DirTree1 
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9551
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DirTree1.LoadTreeDir True, True
End Sub
Private Sub Command2_Click()
DirTree1.LoadTreeDir True
End Sub
Private Sub Command3_Click()
DirTree1.LoadTreeDir True, , , True, True
End Sub
Private Sub Command4_Click()
DirTree1.LoadTreeDir True, , True
End Sub
Private Sub Form_Load()
DirTree1.LoadTreeDir
End Sub
