VERSION 5.00
Begin VB.Form Fsplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2205
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Fsplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ObAV.ucFrame ucFrame1 
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      Caption         =   "obAV Guard"
      Begin VB.Label Label1 
         Caption         =   "Starting Engine"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Creating Tray Icon"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Starting obAV Hook"
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
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Starting IE Watch"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   2
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   1
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   0
         Top             =   1440
         Width           =   255
      End
   End
End
Attribute VB_Name = "Fsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Sub Form_Load()
Me.Caption = vbNullChar
onTOP hwnd
FormCenter Me
End Sub
Sub starTT()
On Error GoTo errH
Dim i As Long
Show
'If App.PrevInstance Then
If FindWindow("#32770", "obav_22091993") > 1 Then
MsgBox "ObAV Guard is Running", vbSystemModal: End
End If
For i = 0 To 4
Select Case i
    Case 0
    Case 1
    bsTray.IconTooltip fRtpSystem.Icon, verobAV
    bsTray.ShellTrayIconAdd fRtpSystem.hwnd
    Case 2
    Load fRtpSystem
    Case 3
    Call cWINHOOK
    Case 4
    fRtpSystem.GuardSTART
End Select
Label2(i).Caption = "ok"
errH:
DoEvents: Sleep 500
Next i
bsTray.ShellTrayBalloonTipShow fRtpSystem.hwnd, 1, verobAV, "Komputer Anda Di proteksi oleh ObAV Guard" & vbCrLf & _
"Oblank AntiVirus System Guard"
Unload Me
End Sub
