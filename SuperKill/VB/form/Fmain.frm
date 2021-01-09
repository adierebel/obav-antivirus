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
   Begin VB.Frame Frame1 
      Caption         =   "Super proccess killer"
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "super kill"
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "process id :"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If Len(Text1.Text) > 1 Then
KeKillProcess Text1.Text
Else
MsgBox "isi dengan proccessID", vbCritical + vbSystemModal
End If
End Sub
Private Sub Form_Load()
App.Title = vbNullChar
BuildServis
End Sub
Private Sub Form_Unload(Cancel As Integer)
ClirServis
End Sub
