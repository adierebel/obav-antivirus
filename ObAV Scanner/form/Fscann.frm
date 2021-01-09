VERSION 5.00
Begin VB.Form Fscann 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6750
   ClientLeft      =   3000
   ClientTop       =   1890
   ClientWidth     =   9360
   Icon            =   "Fscann.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Fscann.frx":591A
   ScaleHeight     =   6750
   ScaleWidth      =   9360
   Begin VB.PictureBox menu1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      MouseIcon       =   "Fscann.frx":BC20
      MousePointer    =   99  'Custom
      Picture         =   "Fscann.frx":BD72
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   73
      Top             =   1440
      Width           =   1815
      Begin VB.Label menus1 
         BackStyle       =   0  'Transparent
         Caption         =   "Protection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MouseIcon       =   "Fscann.frx":E51C
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   75
         Width           =   1215
      End
      Begin VB.Image up1 
         Height          =   120
         Left            =   1560
         Picture         =   "Fscann.frx":E66E
         Top             =   120
         Width           =   150
      End
      Begin VB.Image down1 
         Height          =   120
         Left            =   1560
         Picture         =   "Fscann.frx":E729
         Top             =   120
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.PictureBox menu2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      MouseIcon       =   "Fscann.frx":E7E3
      MousePointer    =   99  'Custom
      Picture         =   "Fscann.frx":E935
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   71
      Top             =   2760
      Width           =   1815
      Begin VB.Label menus2 
         BackStyle       =   0  'Transparent
         Caption         =   "Administration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MouseIcon       =   "Fscann.frx":110DF
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   75
         Width           =   1455
      End
      Begin VB.Image up2 
         Height          =   120
         Left            =   1560
         Picture         =   "Fscann.frx":11231
         Top             =   120
         Width           =   150
      End
      Begin VB.Image down2 
         Height          =   120
         Left            =   1560
         Picture         =   "Fscann.frx":112EC
         Top             =   120
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.PictureBox menu3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      MouseIcon       =   "Fscann.frx":113A6
      MousePointer    =   99  'Custom
      Picture         =   "Fscann.frx":114F8
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   69
      Top             =   5760
      Width           =   1815
      Begin VB.Image Up3 
         Height          =   120
         Left            =   1560
         Picture         =   "Fscann.frx":13CA2
         Top             =   120
         Width           =   150
      End
      Begin VB.Image dOwn3 
         Height          =   120
         Left            =   1560
         Picture         =   "Fscann.frx":13D5D
         Top             =   120
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label menus3 
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MouseIcon       =   "Fscann.frx":13E17
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   80
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      Picture         =   "Fscann.frx":13F69
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   67
      Top             =   6240
      Width           =   1815
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         MouseIcon       =   "Fscann.frx":16713
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   70
         Width           =   1815
      End
   End
   Begin VB.PictureBox MachineTimer 
      Height          =   855
      Left            =   720
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   66
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
      Begin VB.Timer TimerHelp 
         Enabled         =   0   'False
         Index           =   1
         Interval        =   1
         Left            =   0
         Top             =   360
      End
      Begin VB.Timer TimerHelp 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer TmrMenu2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   360
         Top             =   0
      End
      Begin VB.Timer TmrMenu1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   360
         Top             =   360
      End
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   300
      Left            =   960
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   0
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   1440
   End
   Begin ObAVScanner.jcbutton TAB 
      Height          =   735
      Index           =   0
      Left            =   2520
      TabIndex        =   20
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "Scanner"
      MousePointer    =   99
      MouseIcon       =   "Fscann.frx":16865
      PictureNormal   =   "Fscann.frx":169C7
      PictureAlign    =   0
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      DisabledPictureMode=   1
      CaptionEffects  =   1
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin ObAVScanner.jcbutton TAB 
      Height          =   735
      Index           =   1
      Left            =   2520
      TabIndex        =   21
      Top             =   9720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "Guard"
      MousePointer    =   99
      MouseIcon       =   "Fscann.frx":172A1
      PictureNormal   =   "Fscann.frx":17403
      PictureAlign    =   0
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   1
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin ObAVScanner.jcbutton TAB 
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   8280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "Setting"
      MousePointer    =   99
      MouseIcon       =   "Fscann.frx":17CDD
      PictureNormal   =   "Fscann.frx":17E3F
      PictureAlign    =   0
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   1
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin ObAVScanner.jcbutton TAB 
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   9120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "Tool 's"
      MousePointer    =   99
      MouseIcon       =   "Fscann.frx":18719
      PictureNormal   =   "Fscann.frx":1887B
      PictureAlign    =   0
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   1
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin ObAVScanner.jcbutton TAB 
      Height          =   735
      Index           =   4
      Left            =   2520
      TabIndex        =   24
      Top             =   8880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "Quarantine"
      MousePointer    =   99
      MouseIcon       =   "Fscann.frx":19155
      PictureNormal   =   "Fscann.frx":192B7
      PictureAlign    =   0
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   1
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin ObAVScanner.jcbutton TAB 
      Height          =   465
      Index           =   6
      Left            =   1320
      TabIndex        =   25
      Top             =   6960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   820
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "Exit"
      MousePointer    =   99
      MouseIcon       =   "Fscann.frx":19B91
      PictureNormal   =   "Fscann.frx":19CF3
      PictureAlign    =   0
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   1
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin ObAVScanner.jcbutton TAB 
      Height          =   495
      Index           =   5
      Left            =   9960
      TabIndex        =   31
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421504
      Caption         =   "About"
      ForeColor       =   16711680
      MousePointer    =   99
      MouseIcon       =   "Fscann.frx":1A28D
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   0
      Picture         =   "Fscann.frx":1A3EF
      ScaleHeight     =   1320
      ScaleWidth      =   9375
      TabIndex        =   19
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   620
         Left            =   7800
         Picture         =   "Fscann.frx":237FB
         ScaleHeight     =   585
         ScaleWidth      =   1440
         TabIndex        =   38
         Top             =   480
         Width           =   1470
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Indonesia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   40
            Top             =   350
            Width           =   1455
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Made In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   40
            Width           =   1455
         End
      End
      Begin VB.PictureBox secureBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         Picture         =   "Fscann.frx":25F04
         ScaleHeight     =   1335
         ScaleWidth      =   5535
         TabIndex        =   107
         Top             =   0
         Width           =   5535
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver.1.4 Final"
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
            Left            =   3720
            TabIndex        =   109
            Top             =   165
            Width           =   1935
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "19 Oct 2012"
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
            Left            =   3720
            TabIndex        =   108
            Top             =   385
            Width           =   2175
         End
      End
      Begin VB.PictureBox dangerBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         Picture         =   "Fscann.frx":2F250
         ScaleHeight     =   1335
         ScaleWidth      =   5535
         TabIndex        =   110
         Top             =   0
         Visible         =   0   'False
         Width           =   5535
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "23 nov 2012"
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
            Left            =   3720
            TabIndex        =   112
            Top             =   385
            Width           =   2175
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver.1.4 Final"
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
            Left            =   3720
            TabIndex        =   111
            Top             =   165
            Width           =   1935
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   """Local AntiVirus With Heuristic"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   32
         Top             =   165
         Width           =   3855
      End
   End
   Begin VB.PictureBox MeNUL1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   135
      Picture         =   "Fscann.frx":38253
      ScaleHeight     =   1215
      ScaleWidth      =   1800
      TabIndex        =   75
      Top             =   1440
      Width           =   1800
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Guard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         MouseIcon       =   "Fscann.frx":3AF26
         MousePointer    =   99  'Custom
         TabIndex        =   78
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   470
         MouseIcon       =   "Fscann.frx":3B078
         MousePointer    =   99  'Custom
         TabIndex        =   77
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Scanner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   76
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   120
         Picture         =   "Fscann.frx":3B1CA
         Top             =   480
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   120
         Picture         =   "Fscann.frx":3B754
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.PictureBox Menul2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   135
      Picture         =   "Fscann.frx":3BCDE
      ScaleHeight     =   1575
      ScaleWidth      =   1800
      TabIndex        =   79
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Quarantine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         MouseIcon       =   "Fscann.frx":3EC49
         MousePointer    =   99  'Custom
         TabIndex        =   82
         Top             =   480
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "Fscann.frx":3ED9B
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         MouseIcon       =   "Fscann.frx":3F325
         MousePointer    =   99  'Custom
         TabIndex        =   81
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   120
         Picture         =   "Fscann.frx":3F477
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   120
         Picture         =   "Fscann.frx":3FA01
         Top             =   825
         Width           =   240
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Tools"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         MouseIcon       =   "Fscann.frx":3FF8B
         MousePointer    =   99  'Custom
         TabIndex        =   80
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.PictureBox MeNUl3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   135
      Picture         =   "Fscann.frx":400DD
      ScaleHeight     =   1215
      ScaleWidth      =   1800
      TabIndex        =   83
      Top             =   4920
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Image Image7 
         Height          =   240
         Left            =   120
         Picture         =   "Fscann.frx":42DB0
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   86
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1123 
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   470
         MouseIcon       =   "Fscann.frx":4333A
         MousePointer    =   99  'Custom
         TabIndex        =   85
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "Fscann.frx":4348C
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         MouseIcon       =   "Fscann.frx":43A16
         MousePointer    =   99  'Custom
         TabIndex        =   84
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   0
      Left            =   2040
      ScaleHeight     =   5145
      ScaleWidth      =   7185
      TabIndex        =   29
      Top             =   1440
      Width           =   7215
      Begin ObAVScanner.uTabSonny tabson 
         Height          =   4900
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8652
         tabcount        =   5
         aktif           =   2
         judul(1)        =   "Scanner"
         judul(2)        =   "Report"
         judul(3)        =   "Virus @ 0"
         judul(4)        =   "Registry @ 0"
         judul(5)        =   "Hidden @ 0"
         Begin VB.CommandButton Command1 
            Caption         =   "Scan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   10240
            MouseIcon       =   "Fscann.frx":43B68
            MousePointer    =   99  'Custom
            TabIndex        =   65
            Top             =   4320
            Width           =   1575
         End
         Begin ObAVScanner.DirTree DirTree1 
            Height          =   3615
            Left            =   10240
            TabIndex        =   64
            Top             =   600
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   6376
         End
         Begin VB.PictureBox Picture2 
            Height          =   4335
            Index           =   3
            Left            =   -29880
            ScaleHeight     =   4275
            ScaleWidth      =   6675
            TabIndex        =   60
            Top             =   480
            Width           =   6735
            Begin VB.CommandButton Command2 
               Caption         =   "Unhide Check"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   120
               MouseIcon       =   "Fscann.frx":43CBA
               MousePointer    =   99  'Custom
               TabIndex        =   62
               Top             =   3840
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Select All"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   61
               Top             =   70
               Width           =   1095
            End
            Begin ObAVScanner.ucListView ucListView1 
               Height          =   3375
               Index           =   2
               Left            =   120
               TabIndex        =   63
               Top             =   360
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   5953
               StyleEx         =   37
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   4335
            Index           =   2
            Left            =   -19880
            ScaleHeight     =   4275
            ScaleWidth      =   6675
            TabIndex        =   56
            Top             =   480
            Width           =   6735
            Begin VB.CommandButton Command2 
               Caption         =   "Repair Check"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   120
               MouseIcon       =   "Fscann.frx":43E0C
               MousePointer    =   99  'Custom
               TabIndex        =   58
               Top             =   3840
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Select All"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   57
               Top             =   70
               Width           =   1095
            End
            Begin ObAVScanner.ucListView ucListView1 
               Height          =   3375
               Index           =   1
               Left            =   120
               TabIndex        =   59
               Top             =   360
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   5953
               StyleEx         =   37
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   4335
            Index           =   1
            Left            =   -9880
            ScaleHeight     =   4275
            ScaleWidth      =   6675
            TabIndex        =   52
            Top             =   480
            Width           =   6735
            Begin VB.CommandButton Command2 
               Caption         =   "Clean Check"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   120
               MouseIcon       =   "Fscann.frx":43F5E
               MousePointer    =   99  'Custom
               TabIndex        =   54
               Top             =   3840
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Select All"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   53
               Top             =   70
               Width           =   1095
            End
            Begin ObAVScanner.ucListView ucListView1 
               Height          =   3375
               Index           =   0
               Left            =   120
               TabIndex        =   55
               Top             =   360
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   5953
               StyleEx         =   37
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   4095
            Index           =   0
            Left            =   240
            ScaleHeight     =   4035
            ScaleWidth      =   6435
            TabIndex        =   42
            Top             =   600
            Width           =   6495
            Begin VB.CommandButton Command1 
               Caption         =   "Result Virus"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Index           =   1
               Left            =   3600
               MouseIcon       =   "Fscann.frx":440B0
               MousePointer    =   99  'Custom
               TabIndex        =   43
               Top             =   2520
               Width           =   2655
            End
            Begin ObAVScanner.ucProgressBar progBSC 
               Height          =   255
               Left            =   240
               Top             =   1800
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   450
            End
            Begin ObAVScanner.ucFrame ucFrame1 
               Height          =   1575
               Index           =   1
               Left            =   240
               Top             =   2280
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   2778
               Begin VB.Label Label1 
                  Caption         =   "File Scanned       :"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   49
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   255
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   48
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Object Detected  :"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   47
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   255
                  Index           =   1
                  Left            =   1800
                  TabIndex        =   46
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "File Total              :"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   45
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   255
                  Index           =   2
                  Left            =   1800
                  TabIndex        =   44
                  Top             =   1080
                  Width           =   975
               End
            End
            Begin ObAVScanner.ucFrame ucFrame1 
               Height          =   1335
               Index           =   0
               Left            =   240
               Top             =   240
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   2355
               Begin VB.TextBox Text1 
                  Height          =   615
                  Left            =   240
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   50
                  Top             =   480
                  Width           =   5535
               End
               Begin VB.Label Label1 
                  Caption         =   "Path Scan     :"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   51
                  Top             =   240
                  Width           =   975
               End
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   5
      Left            =   2040
      ScaleHeight     =   5145
      ScaleWidth      =   7185
      TabIndex        =   26
      Top             =   1440
      Width           =   7215
      Begin VB.PictureBox Picture7 
         Height          =   4695
         Index           =   1
         Left            =   240
         ScaleHeight     =   4635
         ScaleWidth      =   6675
         TabIndex        =   27
         Top             =   240
         Width           =   6735
         Begin VB.Frame Frame5 
            Height          =   495
            Left            =   120
            TabIndex        =   104
            Top             =   4080
            Width           =   4095
            Begin VB.Label Label28 
               Caption         =   "Click Here"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2520
               MouseIcon       =   "Fscann.frx":44202
               MousePointer    =   99  'Custom
               TabIndex        =   106
               ToolTipText     =   "Click to View"
               Top             =   195
               Width           =   1095
            End
            Begin VB.Label Label27 
               Caption         =   "Please, Upload your sample virus"
               Height          =   255
               Left            =   120
               TabIndex        =   105
               Top             =   195
               Width           =   2415
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "ObAV Guard"
            Height          =   1095
            Left            =   120
            TabIndex        =   94
            Top             =   1680
            Width           =   4095
            Begin VB.Timer Timer3 
               Interval        =   100
               Left            =   0
               Top             =   0
            End
            Begin VB.Label Label23 
               Caption         =   "-"
               Height          =   255
               Left            =   2520
               TabIndex        =   100
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label22 
               Caption         =   "-"
               Height          =   255
               Left            =   2520
               TabIndex        =   99
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label7 
               Caption         =   "ObAV Explorer Guard"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   96
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label7 
               Caption         =   "ObAV EXE Guard"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   95
               Top             =   720
               Width           =   1815
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "General Information"
            Height          =   1455
            Left            =   120
            TabIndex        =   89
            Top             =   120
            Width           =   4095
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   91
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label Label21 
               Caption         =   ":"
               Height          =   255
               Left            =   1725
               TabIndex        =   98
               Top             =   1080
               Width           =   2295
            End
            Begin VB.Label Label8 
               Caption         =   "Release Date"
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label Label6 
               Caption         =   ":"
               Height          =   255
               Left            =   1725
               TabIndex        =   93
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label Label5 
               Caption         =   "ObAV Version"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "ObAV Path Instalation :"
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "ObAV Database"
            Height          =   4455
            Left            =   4320
            TabIndex        =   87
            Top             =   120
            Width           =   2200
            Begin ObAVScanner.ucListView ucLvSign 
               Height          =   4095
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   1950
               _ExtentX        =   3625
               _ExtentY        =   7223
               StyleEx         =   32
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Website :"
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   2880
            Width           =   4095
            Begin VB.Label Label26 
               Caption         =   "www.obav.net"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   720
               MouseIcon       =   "Fscann.frx":44354
               MousePointer    =   99  'Custom
               TabIndex        =   103
               ToolTipText     =   "Click to View"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "Like ObAV Antvirus On"
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label24 
               Caption         =   "Visit Me"
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label9 
               Caption         =   "facebook"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   1800
               MouseIcon       =   "Fscann.frx":444A6
               MousePointer    =   99  'Custom
               TabIndex        =   30
               ToolTipText     =   "Click to View"
               Top             =   720
               Width           =   855
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   2
      Left            =   2040
      ScaleHeight     =   5145
      ScaleWidth      =   7185
      TabIndex        =   18
      Top             =   1440
      Width           =   7215
      Begin ObAVScanner.uTabSonny uTabSonny123 
         Height          =   4905
         Left            =   120
         TabIndex        =   113
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8652
         tabcount        =   2
         judul(1)        =   "General Setting"
         judul(2)        =   "Scan Setting"
         Begin VB.PictureBox Picture222 
            Height          =   4335
            Left            =   -9880
            ScaleHeight     =   4275
            ScaleWidth      =   6675
            TabIndex        =   122
            Top             =   435
            Width           =   6735
            Begin ObAVScanner.ucFrame ucFrame3ff 
               Height          =   735
               Left            =   240
               Top             =   3360
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   1296
               Caption         =   "Info"
               Begin VB.Label LblInfop 
                  Caption         =   "Setting for Scanning file process."
                  Height          =   255
                  Left            =   120
                  TabIndex        =   123
                  Top             =   320
                  Width           =   6015
               End
            End
            Begin ObAVScanner.ucFrame ucFrame1ffcxv 
               Height          =   3015
               Index           =   4
               Left            =   240
               Top             =   240
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   5318
               Caption         =   "Scan Setting's"
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Heuristic String (Alman, Doc Infected)"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   130
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   3735
               End
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Heuristic PE Header (Sality, Virus, etc)"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   129
                  Top             =   720
                  Value           =   1  'Checked
                  Width           =   3735
               End
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Heuristic Icon"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   128
                  Top             =   1080
                  Value           =   1  'Checked
                  Width           =   3495
               End
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Suspect VMX Extention (Like Conficker)"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   127
                  Top             =   1440
                  Value           =   1  'Checked
                  Width           =   4095
               End
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Heuristic Version Header (Fake File)"
                  Height          =   255
                  Index           =   5
                  Left            =   240
                  TabIndex        =   126
                  Top             =   1800
                  Value           =   1  'Checked
                  Width           =   3495
               End
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Heuristic MalScript/VBS"
                  Height          =   255
                  Index           =   6
                  Left            =   240
                  TabIndex        =   125
                  Top             =   2160
                  Value           =   1  'Checked
                  Width           =   3495
               End
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Heuristic Sortcut (Yuyun, Serviks, etc)"
                  Height          =   255
                  Index           =   7
                  Left            =   240
                  TabIndex        =   124
                  Top             =   2520
                  Value           =   1  'Checked
                  Width           =   3735
               End
            End
         End
         Begin VB.PictureBox Pictur22e1 
            Height          =   4335
            Left            =   120
            ScaleHeight     =   4275
            ScaleWidth      =   6675
            TabIndex        =   114
            Top             =   435
            Width           =   6735
            Begin ObAVScanner.ucFrame ucFrame2sd 
               Height          =   1455
               Left            =   240
               Top             =   240
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   2566
               Caption         =   "Window Setting's"
               Begin VB.CheckBox chkOP 
                  Caption         =   "Anti Destroy Window"
                  Height          =   255
                  Index           =   9
                  Left            =   240
                  TabIndex        =   117
                  Top             =   360
                  Width           =   5895
               End
               Begin VB.CheckBox chkOP 
                  Caption         =   "Show On top Window"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   116
                  Top             =   720
                  Width           =   5895
               End
               Begin VB.CheckBox StArTup 
                  Caption         =   "Show Startup Screen"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   115
                  Top             =   1080
                  Width           =   5895
               End
            End
            Begin ObAVScanner.ucFrame ucFrame1fgh 
               Height          =   1455
               Left            =   240
               Top             =   1800
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   2566
               Caption         =   "Extra Setting's"
               Begin VB.CheckBox chkOP 
                  Caption         =   "Enable Scan with in Explorer"
                  Height          =   255
                  Index           =   8
                  Left            =   240
                  TabIndex        =   120
                  Top             =   360
                  Width           =   5895
               End
               Begin VB.CheckBox ScanFD 
                  Caption         =   "Automatic Scan Removable Disk"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   119
                  Top             =   1080
                  Width           =   5895
               End
               Begin VB.CheckBox snd 
                  Caption         =   "Sound Warning"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   118
                  Top             =   720
                  Width           =   5895
               End
            End
            Begin ObAVScanner.ucFrame ucFrame1 
               Height          =   735
               Index           =   4
               Left            =   240
               Top             =   3360
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   1296
               Caption         =   "Info"
               Begin VB.Label infolbl 
                  Caption         =   "General setting for ObAV"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   121
                  Top             =   320
                  Width           =   6015
               End
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   4
      Left            =   2040
      ScaleHeight     =   5145
      ScaleWidth      =   7185
      TabIndex        =   12
      Top             =   1440
      Width           =   7215
      Begin VB.PictureBox Picture7 
         Height          =   4695
         Index           =   0
         Left            =   240
         ScaleHeight     =   4635
         ScaleWidth      =   6675
         TabIndex        =   13
         Top             =   240
         Width           =   6735
         Begin VB.CommandButton Command4 
            Caption         =   "Restore"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            MouseIcon       =   "Fscann.frx":445F8
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check All"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   50
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1440
            MouseIcon       =   "Fscann.frx":4474A
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   4200
            Width           =   1215
         End
         Begin ObAVScanner.ucListView ucLvQuar 
            Height          =   3720
            Left            =   120
            TabIndex        =   16
            Top             =   375
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   7646
            StyleEx         =   37
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   3
      Left            =   2040
      ScaleHeight     =   5145
      ScaleWidth      =   7185
      TabIndex        =   10
      Top             =   1440
      Width           =   7215
      Begin ObAVScanner.uTabSonny uTabSonny1 
         Height          =   4900
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8652
         tabcount        =   2
         judul(1)        =   "Process Manager"
         judul(2)        =   "Start-up Controls"
         Begin VB.PictureBox Picture6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4335
            Index           =   0
            Left            =   120
            ScaleHeight     =   4275
            ScaleWidth      =   6675
            TabIndex        =   36
            Top             =   480
            Width           =   6735
            Begin ObAVScanner.ucListView ucListView2 
               Height          =   4095
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Top             =   120
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   7223
               Style           =   8
               StyleEx         =   33
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   4335
            Index           =   1
            Left            =   -9880
            ScaleHeight     =   4275
            ScaleWidth      =   6675
            TabIndex        =   34
            Top             =   480
            Width           =   6735
            Begin ObAVScanner.ucListView ucListView2 
               Height          =   4095
               Index           =   1
               Left            =   120
               TabIndex        =   35
               Top             =   120
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   7223
               StyleEx         =   33
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   1
      Left            =   2040
      ScaleHeight     =   5145
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   1440
      Width           =   7215
      Begin VB.PictureBox Picture3 
         Height          =   4695
         Left            =   240
         ScaleHeight     =   4635
         ScaleWidth      =   6675
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         Begin ObAVScanner.ucFrame ucFrame1 
            Height          =   2055
            Index           =   2
            Left            =   240
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3625
            Caption         =   "ObAV EXE Guard"
            Begin VB.CommandButton Command3 
               Caption         =   "Activate"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   240
               MouseIcon       =   "Fscann.frx":4489C
               MousePointer    =   99  'Custom
               TabIndex        =   3
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   0
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   2
               Top             =   960
               Width           =   5775
            End
            Begin VB.Label Label3 
               Caption         =   "Non Active"
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
               Index           =   0
               Left            =   960
               TabIndex        =   8
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Last Scanned :"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   4
               Top             =   720
               Width           =   1215
            End
            Begin VB.Shape sStatus 
               BackColor       =   &H000000FF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   0
               Left            =   240
               Shape           =   3  'Circle
               Top             =   300
               Width           =   495
            End
         End
         Begin ObAVScanner.ucFrame ucFrame1 
            Height          =   2055
            Index           =   3
            Left            =   240
            Top             =   2400
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3625
            Caption         =   "ObAV Explorer Guard"
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   1
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   960
               Width           =   5775
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Activate"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   240
               MouseIcon       =   "Fscann.frx":449EE
               MousePointer    =   99  'Custom
               TabIndex        =   5
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "Non Active"
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
               Index           =   1
               Left            =   960
               TabIndex        =   9
               Top             =   360
               Width           =   1335
            End
            Begin VB.Shape sStatus 
               BackColor       =   &H000000FF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   1
               Left            =   240
               Shape           =   3  'Circle
               Top             =   280
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Last Scanned :"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   7
               Top             =   720
               Width           =   1215
            End
         End
      End
   End
   Begin VB.Menu popproc 
      Caption         =   "popproc"
      Visible         =   0   'False
      Begin VB.Menu pop 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu pop 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu pop 
         Caption         =   "Terminate Process"
         Index           =   2
      End
      Begin VB.Menu pop 
         Caption         =   "Suspend Thread"
         Index           =   3
      End
      Begin VB.Menu pop 
         Caption         =   "Resume Thread"
         Index           =   4
      End
   End
   Begin VB.Menu popstar 
      Caption         =   "popstar"
      Visible         =   0   'False
      Begin VB.Menu pops 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu pops 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu pops 
         Caption         =   "Remove Start-Up"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Fscann"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pip As String
Dim pip1 As String

Dim X As Boolean
Dim rtpAKTIV As Boolean
Dim guardAKTIV As Boolean

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function RegDeleteKeyW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
Private Sub Check1_Click(Index As Integer)
CekALL Check1(Index), ucListView1(Index)
End Sub
Private Sub Check2_Click()
CekALL Check2, ucLvQuar
End Sub
Sub chkOP_Click(Index As Integer)
'On Error Resume Next
Select Case Index
    Case 0
    If chkOP(Index).Value = 1 Then _
    cString = True Else cString = False
    Case 1
    If chkOP(Index).Value = 1 Then _
    cPEhead = True Else cPEhead = False
    Case 3
    If chkOP(Index).Value = 1 Then _
    cIcon = True Else cIcon = False
    Case 4
    If chkOP(Index).Value = 1 Then _
    cExVmx = True Else cExVmx = False
    Case 5
    If chkOP(Index).Value = 1 Then _
    cVerHead = True Else cVerHead = False
    Case 6
    If chkOP(Index).Value = 1 Then _
    cMalScrip = True Else cMalScrip = False
    Case 7
    If chkOP(Index).Value = 1 Then _
    cSortcut = True Else cSortcut = False
    Case 9
    If chkOP(Index).Value = 1 Then _
    cAntidestroy = True Else cAntidestroy = False
    Case 2
    If chkOP(Index).Value = 1 Then _
    onTOP hwnd Else onTOP hwnd, True
    Case 8
    If chkOP(Index).Value = 1 Then
    If isInstalled = False Then
    MsgBox "Error,You Must Install ObAV", vbCritical + vbSystemModal
    chkOP(Index).Value = 0
    Exit Sub
    End If
    RegObavExt True
    Else
    RegObavExt False
    RegDeleteKeyW HKEY_CLASSES_ROOT, StrPtr("Folder\Shell\ObAV\command")
    RegDeleteKeyW HKEY_CLASSES_ROOT, StrPtr("Folder\Shell\ObAV")
    End If
End Select
End Sub
Private Sub Command1_Click(Index As Integer)
Dim i As Long
Select Case Index
    Case 0
     If MeNUl3.Visible = True Then
     menus3_Click
     End If
     prosSCAN
    Case 1
    If Command1(1).Caption = "Stop" Then
    scAn = False
    Else
    tabson.AktifTab = 3
    End If
End Select
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
    Case 0
    Eksekusi "karantina", ucListView1(0)
    Case 1
    exsekusiReg ucListView1(1)
    Case 2
    exsekusiHiden ucListView1(2)
End Select
End Sub
Private Sub Command3_Click(Index As Integer)
If Command3(Index).Caption = "Activate" Then
 If FindWindow("#32770", "obav_22091993") > 1 Then
    If Index = 0 Then guardSTATE.Poke "aktiv"
    If Index = 1 Then rtpSTATE.Poke "aktiv"
    Command3(Index).Caption = "DeActivate"
 Else
 MsgBox "Guard Not Running", vbCritical + vbSystemModal
 End If
Else
If Index = 0 Then guardSTATE.Poke "tidak aktiv"
If Index = 1 Then rtpSTATE.Poke "tidak aktiv"
Command3(Index).Caption = "Activate"
End If
End Sub
Private Sub Command4_Click(Index As Integer)
Select Case Index
    Case 0
    restoreQuar ucLvQuar
    getQuarantin ucLvQuar
    Case 1
    Eksekusi "hapus", ucLvQuar
    getQuarantin ucLvQuar
End Select
End Sub

Private Sub Form_Load()
Me.Caption = vbNullChar
DirTree1.LoadTreeDir True
drawLV
Label10_Click
FormCenter Me
getProces ucListView2(0)
GetSTARTUP ucListView2(1)
getDBNFO ucLvSign
getsettingAPP
menu2.Top = 2760
Menul2.Top = 2760
Label10_Click
infor
End Sub
Private Sub infor()
Text3.Text = App.path
Label6.Caption = ": " & Label17.Caption
Label21.Caption = ": " & Label18.Caption
Label29.Caption = Label17.Caption
Label30.Caption = Label18.Caption
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If cAntidestroy = True Then
If X = False Then Cancel = 1
Else
    If Command1(1).Caption = "Stop" Then
        If MsgBox("Are you sure want to exit when process scanning ?", vbQuestion + vbYesNo) = vbYes Then
        TAB_Click (6)
        Else
        Cancel = 1
        End If
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   savesettingAPP
End Sub
Private Sub Label26_Click()
OpenURL "http://www.obav.net", Me.hwnd
End Sub
Private Sub Label28_Click()
OpenURL "http://www.obav.net", Me.hwnd
End Sub
Private Sub Label4_Click()
OpenURL "http://www.obav.net", Me.hwnd
End Sub
Private Sub Label9_Click()
OpenURL "http://facebook.com/obavantivirus", Me.hwnd
End Sub
Private Sub cekItOUt()
If Command3(1).Caption = "DeActivate" And Command3(0).Caption = "DeActivate" Then 'explo
secureBox.Visible = True
dangerBox.Visible = False
Else
secureBox.Visible = False
dangerBox.Visible = True
End If
End Sub

Private Sub pop_Click(Index As Integer)
Select Case Index
    Case 0
    getProces ucListView2(0)
    Case 2
    exPIDLV ucListView2(0), 1
    Case 3
    exPIDLV ucListView2(0), 2
    Case 4
    exPIDLV ucListView2(0), 3
End Select
End Sub
Private Sub pops_Click(Index As Integer)
Dim lvSUB2 As String
Select Case Index
    Case 0
    GetSTARTUP ucListView2(1)
    Case 2
    lvSUB2 = pip
    If Left$(lvSUB2, 18) = "HKEY_LOCAL_MACHINE" Then
    DelSetting HKEY_LOCAL_MACHINE, Right$(lvSUB2, Len(lvSUB2) - 19), pip1
    ElseIf Left$(lvSUB2, 17) = "HKEY_CURRENT_USER" Then
    DelSetting HKEY_CURRENT_USER, Right$(lvSUB2, Len(lvSUB2) - 18), pip1
    GetSTARTUP ucListView2(1)
    End If
End Select
End Sub

Private Sub TAB_Click(Index As Integer)
Dim i As Long
Select Case Index
    Case 4
    getQuarantin ucLvQuar
    Case 6
    X = True
    'Unload Fengine
    'Unload Me
    savesettingAPP
    ExitProcess 0
End Select

For i = 0 To Me.TAB.UBound - 1
    If i <> Index Then
        Picture1(i).Visible = False
        Me.TAB(i).Value = False
    End If
Next i
    Me.TAB(Index).Value = True
    Picture1(Index).Visible = True
End Sub
Private Sub drawLV()
Dim Col As cColumns
With Me
Set Col = .ucListView2(0).Columns
Col.Add , "Proc Name", , , , 2000
Col.Add , "Proc Path", , , , 5000
Col.Add , "PID", , , , 500

Set Col = .ucListView2(1).Columns
Col.Add , "Start-up Name", , , , 2000
Col.Add , "Start-up Path", , , , 5000
Col.Add , "Registry Path", , , , 2000

Set Col = .ucListView1(0).Columns
Col.Add , "Threat Name", , , , 1700
Col.Add , "Threat Path", , , , 4500
Col.Add , "Type", , , , 1000
Col.Add , "Info", , , , 2500

Set Col = .ucListView1(1).Columns
Col.Add , "Value Name", , , , 1700
Col.Add , "Key Path", , , , 5000
Col.Add , "Type", , , , 2000

Set Col = .ucListView1(2).Columns
Col.Add , "File Name", , , , 2000
Col.Add , "File Path", , , , 5000

Set Col = .ucLvQuar.Columns
Col.Add , "File Name", , , , 2000
Col.Add , "Quarant Path", , , , 3000
Col.Add , "Original Path", , , , 5000

Set Col = .ucLvSign.Columns
Col.Add , "No", , , , 400
Col.Add , "Sign Name", , , , 1100

End With
Set Col = Nothing
End Sub
Sub Timer1_Timer()
On Error GoTo er
Dim i As Long
Label2(0).Caption = scannedF
Label2(1).Caption = jmlVIR
Text1.Text = LokasiD
i = (100 / Perc) * scannedF
If i <= 100 Then
progBSC.Value = i
End If
er:
End Sub
Private Sub Timer2_Timer()
Call cekMEM
Call cekMEMGuard
cmdAktiv guardAKTIV, sStatus(0), Label3(0), Command3(0)
cmdAktiv rtpAKTIV, sStatus(1), Label3(1), Command3(1)

Text2(0).Text = TrimW(guardPATH.Peek)
Text2(1).Text = TrimW(rtpPATH.Peek)
End Sub
Sub cekMEM()
If Trim$(TrimW(rtpSTATE.Peek)) = "aktiv" Then
rtpAKTIV = True
Else
rtpAKTIV = False
End If
End Sub
Sub cekMEMGuard()
If Trim$(TrimW(guardSTATE.Peek)) = "aktiv" Then
guardAKTIV = True
Else
guardAKTIV = False
End If
End Sub
Private Sub Timer3_Timer()
cekItOUt
If Command3(1).Caption = "DeActivate" Then 'explo
Label22.Caption = "Running"
Else
Label22.Caption = "Stoped"
End If
If Command3(0).Caption = "DeActivate" Then 'exe
Label23.Caption = "Running"
Else
Label23.Caption = "Stoped"
End If
End Sub
Private Sub ucListView2_ItemClick(Index As Integer, ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
If iButton = vbccMouseRButton Then
    If Index = 0 Then
    PopupMenu popproc, 0, , , pop(0)
    Else
    pip = oItem.SubItem(3).Text
    pip1 = oItem.SubItem(1).Text
    PopupMenu popstar, 0, , , pops(0)
    End If
End If
End Sub

Private Sub ucListView2_ItemSelect(Index As Integer, ByVal oItem As cListItem, ByVal bSelect As Boolean)
If Index = 0 Then
    If bSelect = True Then
    oItem.Checked = True
    Else
    oItem.Checked = False
    End If
End If
End Sub
'********************************************
' maap boss, ane masih bingung tentang case..
'jadinya seperti ini boss, super panjang...
' tolong di ringkas lagi boss...
'********************************************
Private Sub menus2_Click()
If Command1(1).Caption = "Result Virus" Then
    If Menul2.Visible = False Then
    TmrMenu2.Enabled = True
    cekHelp
    End If
End If
End Sub
Private Sub cekHelp()
If MeNUl3.Visible = True Then
TimerHelp(1).Enabled = True
MeNUl3.Visible = False
Up3.Visible = True
dOwn3.Visible = False
End If
End Sub
Private Sub TimerHelp_Timer(Index As Integer)
Select Case Index
Case 0
If menu3.Top = 4920 Then
TimerHelp(0).Enabled = False
MeNUl3.Visible = True
dOwn3.Visible = True
Up3.Visible = False
Label15_Click
End If
menu3.Top = menu3.Top - 40
Case 1
If menu3.Top = 5760 Then
TimerHelp(1).Enabled = False
End If
menu3.Top = menu3.Top + 40
End Select
End Sub
Private Sub TmrMenu1_Timer()
If menu2.Top = 2720 Then
TmrMenu1.Enabled = False
up1.Visible = True
down1.Visible = False
MeNUL1.Visible = True
Label10_Click
End If
up2.Visible = False
down2.Visible = True
Menul2.Visible = False
Menul2.Top = Menul2.Top + 40
menu2.Top = menu2.Top + 40
End Sub
Private Sub menus1_Click()
    If MeNUL1.Visible = False Then
    TmrMenu1.Enabled = True
    cekHelp
    End If
End Sub
Private Sub TmrMenu2_Timer()
If menu2.Top = 1960 Then
TmrMenu2.Enabled = False
Menul2.Top = 1920
up2.Visible = True
down2.Visible = False
Menul2.Visible = True
Label11_Click
End If
MeNUL1.Visible = False
up1.Visible = False
down1.Visible = True
menu2.Top = menu2.Top - 40
End Sub
Private Sub menu1_Click()
menus1_Click
End Sub
Private Sub menu2_Click()
menus2_Click
End Sub
Private Sub Label10_Click()
TAB_Click (0)
Label10.FontUnderline = True
Label14.FontUnderline = False
Label11.FontUnderline = False
Label12.FontUnderline = False
Label13.FontUnderline = False
Label15.FontUnderline = False
Label1123.FontUnderline = False
End Sub
Private Sub Label11_Click()
TAB_Click (4)
Label10.FontUnderline = False
Label14.FontUnderline = False
Label11.FontUnderline = True
Label12.FontUnderline = False
Label13.FontUnderline = False
Label15.FontUnderline = False
Label1123.FontUnderline = False
End Sub
Private Sub Label12_Click()
TAB_Click (3)
Label10.FontUnderline = False
Label14.FontUnderline = False
Label11.FontUnderline = False
Label12.FontUnderline = True
Label13.FontUnderline = False
Label15.FontUnderline = False
Label1123.FontUnderline = False
End Sub
Private Sub Label13_Click()
TAB_Click (2)
Label10.FontUnderline = False
Label14.FontUnderline = False
Label11.FontUnderline = False
Label12.FontUnderline = False
Label13.FontUnderline = True
Label15.FontUnderline = False
Label1123.FontUnderline = False
End Sub
Private Sub Label14_Click()
If Command1(1).Caption = "Result Virus" Then
    TAB_Click (1)
    Label10.FontUnderline = False
    Label14.FontUnderline = True
    Label11.FontUnderline = False
    Label12.FontUnderline = False
    Label13.FontUnderline = False
    Label15.FontUnderline = False
    Label1123.FontUnderline = False
End If
End Sub
Private Sub Label15_Click()
TAB_Click (5)
Label10.FontUnderline = False
Label14.FontUnderline = False
Label11.FontUnderline = False
Label12.FontUnderline = False
Label13.FontUnderline = False
Label15.FontUnderline = True
Label1123.FontUnderline = False
End Sub
Private Sub Label1123_Click()
Fabout.Show
End Sub
Private Sub Label16_Click()
If Command1(1).Caption = "Stop" Then
    If MsgBox("Are you sure want to exit when process scanning ?", vbQuestion + vbYesNo) = vbYes Then
    TAB_Click (6)
    End If
Else
TAB_Click (6)
End If
End Sub
Private Sub menu3_Click()
menus3_Click
End Sub
Private Sub menus3_Click()
If Command1(1).Caption = "Result Virus" Then
    If MeNUl3.Visible = True Then
    MeNUl3.Visible = False
    Up3.Visible = True
    dOwn3.Visible = False
    TimerHelp(1).Enabled = True
    Else
    TimerHelp(0).Enabled = True
    End If
End If
End Sub
Private Sub chkOP_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
LblInfop.Caption = "Enable or Disable Heuristic String (Alman, Doc Infected)"
Case 1
LblInfop.Caption = "Enable or Disable Heuristic PE Header (Sality, Virus, etc)"
Case 3
LblInfop.Caption = "Enable or Disable Heuristic Icon"
Case 4
LblInfop.Caption = "Enable or Disable Suspect VMX Extention (Like Conficker)"
Case 5
LblInfop.Caption = "Enable or Disable Heuristic Version Header (Fake File)"
Case 6
LblInfop.Caption = "Enable or Disable Heuristic MalScript/VBS"
Case 7
LblInfop.Caption = "Enable or Disable Heuristic Sortcut (Yuyun, Serviks, etc)"
'====
Case 9
infolbl.Caption = "Enable or Disable Anti Destroy Window Mode"
Case 2
infolbl.Caption = "Enable or Disable Show On top Window Mode"
Case 8
infolbl.Caption = "Enable or Disable Shell Scan with in Explorer"
End Select
End Sub
Private Sub Pictur22e1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "General setting for ObAV"
End Sub
Private Sub Picture222_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblInfop.Caption = "Setting for Scanning file process."
End Sub
Private Sub ScanFD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "Enable or Disable Automatic Scan Removable Disk"
End Sub
Private Sub snd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "Enable or Disable Sound Warning"
End Sub
Private Sub StArTup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infolbl.Caption = "Enable or Disable Show Startup Screen"
End Sub


