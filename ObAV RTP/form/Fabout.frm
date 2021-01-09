VERSION 5.00
Begin VB.Form Fabout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ObAV AntiVirus"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Fabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2640
      Left            =   0
      ScaleHeight     =   2640
      ScaleWidth      =   6855
      TabIndex        =   7
      Top             =   2540
      Width           =   6855
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   6000
         Top             =   240
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4005
         Left            =   0
         Picture         =   "Fabout.frx":000C
         ScaleHeight     =   4005
         ScaleWidth      =   5415
         TabIndex        =   8
         Top             =   -3960
         Width           =   5415
         Begin VB.Image Image1 
            Height          =   255
            Index           =   1
            Left            =   120
            MouseIcon       =   "Fabout.frx":10CF8
            MousePointer    =   99  'Custom
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   0
            Left            =   120
            MouseIcon       =   "Fabout.frx":10E4A
            MousePointer    =   99  'Custom
            Top             =   2640
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   0
      Picture         =   "Fabout.frx":10F9C
      ScaleHeight     =   705
      ScaleWidth      =   7575
      TabIndex        =   4
      Top             =   5200
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5550
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "www.obav.net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         MouseIcon       =   "Fabout.frx":14C54
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      Picture         =   "Fabout.frx":14DA6
      ScaleHeight     =   1215
      ScaleWidth      =   7410
      TabIndex        =   3
      Top             =   1320
      Width           =   7410
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   0
      Picture         =   "Fabout.frx":1AEC6
      ScaleHeight     =   1320
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ver.1.1 Alpha"
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
         Left            =   3480
         TabIndex        =   2
         Top             =   130
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "?? ??? 2012"
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
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Fabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
FormCenter Me
onTOP hwnd
Me.Caption = vbNullString
ceksver
End Sub
Private Sub ceksver()
Label17.Caption = Fscann.Label17.Caption
Label18.Caption = Fscann.Label18.Caption
End Sub
Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
OpenURL "http://www.obav.net", Me.hwnd
Case 1
OpenURL "http://facebook.com/obavantivirus", Me.hwnd
End Select
End Sub
Private Sub Label1_Click()
OpenURL "http://www.obav.net", Me.hwnd
End Sub

Private Sub Timer1_Timer()
If Picture4.Top = -3960 Then
Picture4.Top = 2520
Else
Picture4.Top = Picture4.Top - 4
End If
End Sub
