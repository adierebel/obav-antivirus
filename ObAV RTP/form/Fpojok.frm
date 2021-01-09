VERSION 5.00
Begin VB.Form Fpojok 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   6975
   Icon            =   "Fpojok.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Fpojok.frx":000C
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   480
      Top             =   2520
   End
   Begin VB.CommandButton Cstop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin ObAVGuard.ucProgressBar ucProgressBar1 
      Height          =   375
      Left            =   120
      Top             =   1800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
   End
   Begin VB.Label lblDisek 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Nothing"
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
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   900
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   735
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
Dim lonRecT As String
lonRecT = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 10, 10)
SetWindowRgn Me.hwnd, lonRecT, True
SetFormPojok Me
lblDisek.Caption = "Scanning " & fRtpSystem.DRibeLBL.Caption
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        MoveForm Me.hwnd
    End If
End Sub
