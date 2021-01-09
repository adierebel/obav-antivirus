VERSION 5.00
Begin VB.Form Fpojok 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   Icon            =   "Fpojok.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Fpojok.frx":000C
   ScaleHeight     =   1935
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   480
      Top             =   2280
   End
   Begin VB.CommandButton Cstop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin ObAV.ucProgressBar ucProgressBar1 
      Height          =   375
      Left            =   120
      Top             =   1440
      Width           =   5415
      _ExtentX        =   8493
      _ExtentY        =   661
   End
   Begin VB.Label lblDisek 
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   740
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning : "
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Fpojok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cstop_Click()
If Cstop.Caption = "Stop" Then
scAn = False
Cstop.Caption = "Close"
Else
Unload Me
End If
End Sub
Private Sub Form_Load()
SetFormPojok Me
lblDisek.Caption = fRtpSystem.DRibeLBL.Caption
End Sub
Private Sub Timer1_Timer()
On Error GoTo er
Dim i As Long
Text1.Text = LokasiD
i = (100 / Perc) * scannedF
If i <= 100 Then
ucProgressBar1.Value = i
End If
er:
End Sub

