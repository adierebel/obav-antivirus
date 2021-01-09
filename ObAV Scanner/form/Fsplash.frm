VERSION 5.00
Begin VB.Form Fsplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Fsplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   240
      ScaleHeight     =   1500
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   1800
      Width           =   4575
      Begin VB.PictureBox O0kk 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   4320
         Picture         =   "Fsplash.frx":000C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox O0kk 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   4320
         Picture         =   "Fsplash.frx":22CD
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   840
         Width           =   240
      End
      Begin VB.PictureBox O0kk 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   4320
         Picture         =   "Fsplash.frx":458E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   480
         Width           =   240
      End
      Begin VB.PictureBox O0kk 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   4320
         Picture         =   "Fsplash.frx":684F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox XXXX 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   4320
         Picture         =   "Fsplash.frx":8B10
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox XXXX 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   4320
         Picture         =   "Fsplash.frx":AEC8
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   1
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox XXXX 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   4320
         Picture         =   "Fsplash.frx":D280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox XXXX 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   4320
         Picture         =   "Fsplash.frx":F638
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting IE Watch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting ObAV Hook"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creating Tray Icon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Engine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   2655
      End
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   -240
      Picture         =   "Fsplash.frx":119F0
      Top             =   -240
      Width           =   2025
   End
   Begin VB.Image Image2 
      Height          =   4125
      Left            =   -720
      Picture         =   "Fsplash.frx":131C3
      Top             =   -240
      Width           =   6435
   End
End
Attribute VB_Name = "Fsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" Alias "createroundrectrgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" Alias "createrectrgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" Alias "createellipticrgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" Alias "combinergn" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Function fMakeATranspArea(AreaType As String, pCordinate() As Long) As Boolean
Const RGN_DIFF = 4
Dim lOriginalForm As Long
Dim ltheHole As Long
Dim lNewForm As Long
Dim lfWidth As Single
Dim lfHeight As Single
Dim lborder_width As Single
Dim ltitle_height As Single
 On Error GoTo Trap
 lfWidth = ScaleX(Width, vbTwips, vbPixels)
 lfHeight = ScaleY(Height, vbTwips, vbPixels)
 lOriginalForm = CreateRectRgn(0, 0, lfWidth, lfHeight)

 lborder_width = (lfHeight - ScaleWidth) / 2
 ltitle_height = lfHeight - lborder_width - ScaleHeight
Select Case AreaType

 Case "elliptic"

 ltheHole = CreateEllipticRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
 Case "rectangle"

 ltheHole = CreateRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))

 Case "roundrect"

 ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(5), pCordinate(6))
 Case "circle"
 ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(3), pCordinate(4))

 Case Else
 MsgBox "unknown shape!!"
 Exit Function
 End Select
 lNewForm = CreateRectRgn(0, 0, 0, 0)
 CombineRgn lNewForm, lOriginalForm, _
 ltheHole, RGN_DIFF

 SetWindowRgn hwnd, lNewForm, True
 Me.Refresh
 fMakeATranspArea = True
Exit Function
Trap:
 MsgBox "error occurred. error # " & erR.Number & ", " & erR.Description
End Function
Private Sub Form_Load()
Dim i As Integer
    For i = 0 To 255 Step 3
        ActiveTransparency Me, True, False, i
        Me.Refresh
    Next i
Me.Caption = vbNullChar
onTOP hwnd
FormCenter Me
End Sub
Sub starTT()
Dim i As Long
'If App.PrevInstance Then
If FindWindow("#32770", "obav_22091993") > 1 Then
MsgBox "ObAV Guard is Running", vbSystemModal: End
End If
'====================================================
If Fscann.StArTup.Value = 1 Then
Okokok
Me.Visible = True
keLiatanTuh
Else
Me.Visible = False
bsTray.IconTooltip fRtpSystem.Icon, verobAV
bsTray.ShellTrayIconAdd fRtpSystem.hwnd
Load fRtpSystem
Call cWINHOOK
fRtpSystem.GuardSTART
bsTray.ShellTrayBalloonTipShow fRtpSystem.hwnd, 1, verobAV, "Komputer Anda Di proteksi oleh ObAV Guard" & vbCrLf & _
"ObAV AntiVirus System Guard"
Unload Me
End If
End Sub
Private Sub keLiatanTuh()
On Error GoTo errH
Dim i As Long
For i = 0 To 4
Select Case i
    Case 0
    XXXX(0).Visible = False
    O0kk(0).Visible = True
    Case 1
    bsTray.IconTooltip fRtpSystem.Icon, verobAV
    bsTray.ShellTrayIconAdd fRtpSystem.hwnd
    XXXX(1).Visible = False
    O0kk(1).Visible = True
    Case 2
    Load fRtpSystem
    XXXX(2).Visible = False
    O0kk(2).Visible = True
    Case 3
    Call cWINHOOK
    XXXX(3).Visible = False
    O0kk(3).Visible = True
    Case 4
    fRtpSystem.GuardSTART
End Select
Label2(i).Caption = "ok"
errH:
DoEvents: Sleep 500
Next i
Unload Me
End Sub
Private Sub Okokok()
XXXX(0).Visible = True
XXXX(1).Visible = True
XXXX(2).Visible = True
XXXX(3).Visible = True
O0kk(0).Visible = False
O0kk(1).Visible = False
O0kk(2).Visible = False
O0kk(3).Visible = False
End Sub
