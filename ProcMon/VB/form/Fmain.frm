VERSION 5.00
Begin VB.Form Fmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "remnotiv"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "setnotiv"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
SetNotiv
End Sub
Private Sub Command2_Click()
RemovNotiv
End Sub
Private Sub Form_Load()
App.Title = vbNullChar
BuildServis
End Sub
Private Sub Form_Unload(Cancel As Integer)
ClirServis
End Sub
