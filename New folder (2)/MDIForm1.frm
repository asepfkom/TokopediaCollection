VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tele Integrasi Nusantara Sistem Tokopedia Ver 18-06-2020"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   14445
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "MDIForm1.frx":08CA
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerAutoDialer 
      Interval        =   1000
      Left            =   6480
      Top             =   960
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   960
   End
   Begin VB.Timer Timer11 
      Interval        =   5000
      Left            =   5520
      Top             =   960
   End
   Begin VB.Timer Timer10 
      Interval        =   1000
      Left            =   5040
      Top             =   960
   End
   Begin Threed.SSPanel SSPanel6 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   53
      _Version        =   196610
      Caption         =   "SSPanel6"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8880
      Top             =   480
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   6000
      Top             =   0
   End
   Begin VB.Timer TimerRequest 
      Interval        =   40000
      Left            =   7440
      Top             =   0
   End
   Begin VB.Timer TimerTandaReq 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7920
      Top             =   0
   End
   Begin VB.Timer TimerWaktu 
      Interval        =   100
      Left            =   6960
      Top             =   0
   End
   Begin VB.Timer TimerTanda 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5040
      Top             =   480
   End
   Begin VB.Timer Timer_stopwatch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7920
      Top             =   480
   End
   Begin VB.Timer TimerBlink 
      Interval        =   1000
      Left            =   6480
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Interval        =   50000
      Left            =   5535
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   5040
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   8400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9999
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   465
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   480
   End
   Begin VB.Timer TimerCTI 
      Interval        =   300
      Left            =   7440
      Top             =   480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   6960
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   5520
      Top             =   480
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      TabIndex        =   0
      Top             =   7785
      Visible         =   0   'False
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   26
      _Version        =   196610
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox TxtLamaFollowup 
         Height          =   285
         Left            =   1050
         TabIndex        =   10
         Top             =   945
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox TxtJamSelesaiTelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7605
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "23:59:59"
         Top             =   120
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox TxtJamMulaiTelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5715
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   120
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox TxtModemAcod 
         Height          =   285
         Left            =   420
         TabIndex        =   7
         Text            =   "Text8"
         Top             =   1275
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox TxtAuthPrefix 
         Height          =   285
         Left            =   1770
         TabIndex        =   6
         Text            =   "Text8"
         Top             =   2715
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox TxtAuth 
         Height          =   285
         Left            =   4110
         TabIndex        =   5
         Top             =   90
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2700
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   75
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtCommPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2250
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   75
         Width           =   390
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12435
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   75
         Width           =   2685
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Communication"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   285
         TabIndex        =   3
         Top             =   90
         Width           =   1920
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8400
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   6000
   End
   Begin MSCommLib.MSComm MsComLogin 
      Left            =   9360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock WskCTI 
      Left            =   8865
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   18000
   End
   Begin Threed.SSPanel SSPanel_browse 
      Align           =   3  'Align Left
      Height          =   7755
      Left            =   0
      TabIndex        =   58
      Top             =   30
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   13679
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton Command1 
         Caption         =   "<< &Hide"
         Height          =   375
         Left            =   3540
         TabIndex        =   59
         Top             =   60
         Width           =   1335
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   9255
         Left            =   120
         TabIndex        =   60
         Top             =   540
         Width           =   4695
         ExtentX         =   8281
         ExtentY         =   16325
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Align           =   4  'Align Right
      Height          =   7755
      Left            =   14145
      TabIndex        =   67
      Top             =   30
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   13679
      _Version        =   196610
      BackColor       =   14737632
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton Command8 
         Caption         =   "V"
         Height          =   375
         Left            =   600
         TabIndex        =   94
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Set"
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   5280
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   240
         TabIndex        =   90
         Top             =   5880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Edit"
         Height          =   435
         Left            =   120
         TabIndex        =   82
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   80
         Top             =   3360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OFF"
         Height          =   255
         Left            =   45
         TabIndex        =   79
         Top             =   2280
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON"
         Height          =   255
         Left            =   45
         TabIndex        =   78
         Top             =   1920
         Width           =   1050
      End
      Begin VB.CommandButton cmdenabledptp 
         Caption         =   "OK"
         Height          =   315
         Left            =   2370
         MaskColor       =   &H00000000&
         TabIndex        =   74
         Top             =   2700
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "YES  |  NO"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   855
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1508
         _Version        =   196610
         BackColor       =   16711680
         PictureFrames   =   1
         Picture         =   "MDIForm1.frx":2114D
         ButtonStyle     =   2
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hiden Aksesall"
         Height          =   495
         Left            =   120
         TabIndex        =   92
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data per Agent"
         Height          =   495
         Left            =   120
         TabIndex        =   81
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Enabled PTP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         TabIndex        =   75
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ndy - Lite"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Align           =   2  'Align Bottom
      Height          =   1110
      Left            =   0
      TabIndex        =   11
      Top             =   7800
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   1958
      _Version        =   196610
      BackColor       =   12632256
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_f_cp 
         Height          =   315
         Left            =   17910
         TabIndex        =   102
         Top             =   135
         Width           =   855
      End
      Begin VB.TextBox txtuniqueid 
         Height          =   285
         Left            =   21975
         TabIndex        =   101
         Text            =   "uniqueid cti"
         Top             =   75
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtdurasi 
         Height          =   285
         Left            =   21975
         TabIndex        =   100
         Text            =   "durasi"
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin Threed.SSFrame Frame_autodialer 
         Height          =   840
         Left            =   14175
         TabIndex        =   95
         Top             =   330
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   1482
         _Version        =   196610
         BackColor       =   12632256
         Caption         =   "Autodialer"
         Begin Threed.SSCommand Cmdautodialer 
            Height          =   435
            Index           =   0
            Left            =   180
            TabIndex        =   96
            Top             =   270
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   767
            _Version        =   196610
            Caption         =   "Start"
         End
         Begin Threed.SSCommand Cmdautodialer 
            Height          =   435
            Index           =   1
            Left            =   180
            TabIndex        =   97
            Top             =   270
            Visible         =   0   'False
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   767
            _Version        =   196610
            Caption         =   "Stop"
         End
         Begin Threed.SSCommand Cmdautodialer 
            Height          =   435
            Index           =   2
            Left            =   885
            TabIndex        =   98
            Top             =   270
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   767
            _Version        =   196610
            Caption         =   "STOP / Manual Dial"
         End
         Begin VB.Timer Timer_Durasi_status_autodialer 
            Interval        =   1000
            Left            =   1770
            Top             =   240
         End
         Begin VB.Label lblautdialer_timer_start_stop 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2595
            TabIndex        =   99
            Top             =   345
            Width           =   330
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   465
         Left            =   13335
         TabIndex        =   84
         Top             =   1500
         Visible         =   0   'False
         Width           =   945
         Begin Threed.SSCommand SSCommand1 
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   85
            ToolTipText     =   "Find.."
            Top             =   840
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            _Version        =   196610
            Font3D          =   1
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   -2147483638
            PictureMaskColor=   255
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "MDIForm1.frx":22255
            AutoSize        =   1
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   0
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   86
            ToolTipText     =   "Find.."
            Top             =   720
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            _Version        =   196610
            Font3D          =   1
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   -2147483638
            PictureMaskColor=   255
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "MDIForm1.frx":22742
            AutoSize        =   1
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   0
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   180
            Index           =   4
            Left            =   240
            TabIndex        =   87
            ToolTipText     =   "Find.."
            Top             =   600
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            _Version        =   196610
            Font3D          =   1
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   -2147483638
            PictureMaskColor=   255
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "MDIForm1.frx":22C2F
            AutoSize        =   1
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   0
            BevelWidth      =   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   180
            Index           =   5
            Left            =   240
            TabIndex        =   88
            ToolTipText     =   "Find.."
            Top             =   360
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            _Version        =   196610
            Font3D          =   1
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   -2147483638
            PictureMaskColor=   255
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "MDIForm1.frx":2311C
            AutoSize        =   1
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   0
            BevelWidth      =   1
         End
         Begin VB.Label LblTarget 
            Caption         =   "Label4"
            Height          =   135
            Left            =   720
            TabIndex        =   89
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Enabled         =   0   'False
         Height          =   195
         Left            =   19275
         TabIndex        =   77
         Top             =   195
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   13320
         TabIndex        =   76
         Top             =   2265
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   19500
         TabIndex        =   72
         Text            =   "Text4"
         Top             =   390
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CommandButton apaja 
         Caption         =   "Command3"
         Enabled         =   0   'False
         Height          =   195
         Left            =   19035
         TabIndex        =   71
         Top             =   195
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox RichTextBox1 
         Height          =   285
         Left            =   19230
         TabIndex        =   66
         Text            =   "Text4"
         Top             =   390
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CommandButton cmd_break 
         BackColor       =   &H008080FF&
         Caption         =   "Break Time !!"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8445
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   195
         Width           =   2055
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   22410
         Top             =   675
      End
      Begin VB.TextBox TxtIPIcentra 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   21030
         TabIndex        =   63
         Top             =   555
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<< &Show Media"
         Height          =   375
         Left            =   21030
         TabIndex        =   61
         Top             =   195
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSWinsockLib.Winsock WskRequest 
         Index           =   0
         Left            =   20385
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox TxtOnline 
         Height          =   1005
         Left            =   20820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   75
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton CmdInbox 
         Caption         =   "Command1"
         Height          =   195
         Left            =   18810
         TabIndex        =   54
         Top             =   195
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TxtWaktuRefresh 
         Enabled         =   0   'False
         Height          =   285
         Left            =   19005
         TabIndex        =   52
         Text            =   "00:00:30"
         Top             =   765
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3795
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   735
         Width           =   4095
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   510
         Left            =   14355
         TabIndex        =   40
         Top             =   1500
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   900
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "MDIForm1.frx":23609
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "MDIForm1.frx":23625
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "MDIForm1.frx":23641
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10455
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   765
         Width           =   1110
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9225
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   765
         Width           =   1320
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   11490
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   765
         Width           =   2655
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   3105
         Left            =   5640
         TabIndex        =   12
         Top             =   3120
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   5477
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand SSCommand2 
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2880
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   196610
            Caption         =   "&REFRESH"
         End
         Begin MSComctlLib.ListView LstGrade 
            Height          =   3345
            Left            =   -75
            TabIndex        =   14
            Top             =   -390
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   5900
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16761087
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   3000
            Width           =   2295
         End
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   315
         Left            =   8040
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   765
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         Calendar        =   "MDIForm1.frx":2365D
         Caption         =   "MDIForm1.frx":23775
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "MDIForm1.frx":237E1
         Keys            =   "MDIForm1.frx":237FF
         Spin            =   "MDIForm1.frx":2385D
         AlignHorizontal =   0
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   12648447
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   8388608
         Format          =   "dd/mm/yyyy"
         HighlightText   =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   1
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37475
         CenturyMode     =   0
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   450
         Left            =   3810
         TabIndex        =   20
         Top             =   180
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   794
         _Version        =   196610
         BackColor       =   14737632
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox CmbNo 
            Height          =   315
            Left            =   150
            TabIndex        =   21
            Top             =   1710
            Width           =   1590
         End
         Begin Threed.SSCommand CmdCancel 
            Cancel          =   -1  'True
            Height          =   360
            Left            =   1830
            TabIndex        =   22
            Top             =   1605
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   635
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Clear"
            ButtonStyle     =   2
            BevelWidth      =   1
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   2
            Left            =   630
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1410
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "2"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   1
            Left            =   105
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1410
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "1"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   3
            Left            =   1140
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1410
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "3"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   5
            Left            =   2160
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1410
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "5"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   4
            Left            =   1650
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1410
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "4"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   6
            Left            =   120
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1935
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "6"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   8
            Left            =   1140
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1935
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "8"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   7
            Left            =   630
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1935
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "7"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   9
            Left            =   1650
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1935
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "9"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdNo 
            Height          =   480
            Index           =   0
            Left            =   2175
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   1935
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "0"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdBintang 
            Height          =   480
            Left            =   2685
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1410
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "*"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdPager 
            Height          =   480
            Left            =   2700
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1935
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   192
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "#"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdHangUp 
            Height          =   555
            Left            =   3225
            TabIndex        =   35
            ToolTipText     =   "HangUp"
            Top             =   1935
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   979
            _Version        =   196610
            MousePointer    =   16
            PictureFrames   =   1
            Picture         =   "MDIForm1.frx":23885
            Caption         =   "Hangup"
            Alignment       =   8
            ButtonStyle     =   4
            PictureAlignment=   11
            BevelWidth      =   0
         End
         Begin MSComctlLib.ListView LstInformation 
            Height          =   1020
            Left            =   6240
            TabIndex        =   37
            Top             =   1380
            Width           =   4350
            _ExtentX        =   7673
            _ExtentY        =   1799
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin Threed.SSCommand CmdCall 
            Height          =   540
            Left            =   3225
            TabIndex        =   39
            ToolTipText     =   "Call"
            Top             =   1380
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   953
            _Version        =   196610
            MousePointer    =   16
            PictureFrames   =   1
            Picture         =   "MDIForm1.frx":23B0A
            Caption         =   "Call"
            Alignment       =   8
            ButtonStyle     =   4
            PictureAlignment=   11
            BevelWidth      =   0
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00B1FDD5&
            Caption         =   "."
            Height          =   75
            Left            =   4320
            TabIndex        =   70
            Top             =   120
            Width           =   45
         End
         Begin VB.Label LblJmlSmsBaru 
            BackColor       =   &H00E0E0E0&
            Caption         =   "0"
            Height          =   195
            Left            =   300
            TabIndex        =   55
            Top             =   1515
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   660
         Index           =   0
         Left            =   240
         TabIndex        =   38
         ToolTipText     =   "Find.."
         Top             =   120
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1164
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   192
         BackColor       =   -2147483638
         PictureMaskColor=   255
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":23F91
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   0
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   660
         Index           =   8
         Left            =   1080
         TabIndex        =   42
         ToolTipText     =   "Blok Data..."
         Top             =   120
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1164
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   192
         BackColor       =   -2147483638
         PictureMaskColor=   255
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":2447E
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   660
         Index           =   10
         Left            =   1920
         TabIndex        =   43
         ToolTipText     =   "Blok Data..."
         Top             =   120
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1164
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   192
         BackColor       =   -2147483638
         PictureMaskColor=   255
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":24A04
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   660
         Index           =   11
         Left            =   2790
         TabIndex        =   44
         ToolTipText     =   "Blok Data..."
         Top             =   120
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1164
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   192
         BackColor       =   -2147483638
         PictureMaskColor=   255
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":24F45
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin VB.Label Lbltargetspv 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   2175
         Left            =   5970
         TabIndex        =   49
         Top             =   1410
         Width           =   11265
      End
      Begin VB.Label Label_OL_count 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   17010
         TabIndex        =   65
         Top             =   450
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label LblBersihkan 
         BackColor       =   &H80000008&
         Height          =   195
         Left            =   22260
         TabIndex        =   62
         Top             =   825
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape ShapeReq 
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   20010
         Shape           =   3  'Circle
         Top             =   60
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbl_timer_activity 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   21030
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape ShapeTanda 
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8040
         Shape           =   3  'Circle
         Top             =   165
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Label Waktu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15960
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Label Label9 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8100
         TabIndex        =   51
         Top             =   135
         Width           =   4515
      End
      Begin VB.Label Label10 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   12885
         TabIndex        =   50
         Top             =   150
         Width           =   4605
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2850
         TabIndex        =   48
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Broadcast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1860
         TabIndex        =   47
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Block Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1005
         TabIndex        =   46
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Follow Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   195
         TabIndex        =   45
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   270
         Left            =   17895
         TabIndex        =   36
         Top             =   495
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu MnFile 
         Caption         =   "&Log Off"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MnFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnFile 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MnFile 
         Caption         =   "&Status TSE For Spv"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu MnFile 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu MnFile 
         Caption         =   "&Change Password"
         Index           =   5
      End
      Begin VB.Menu MnFile 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu MnFile 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Master"
      Index           =   1
      Begin VB.Menu nmupload 
         Caption         =   "&Upload data"
         Visible         =   0   'False
         Begin VB.Menu nmuploadcustomer 
            Caption         =   "Upload data customer"
            Visible         =   0   'False
         End
         Begin VB.Menu nmuploadpayment 
            Caption         =   "Upload payment"
         End
         Begin VB.Menu mnuuploadcpa 
            Caption         =   "Upload CPA"
            Visible         =   0   'False
         End
         Begin VB.Menu nmswapdata 
            Caption         =   "Swap data"
         End
         Begin VB.Menu nmgupload 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu nmrestoredeleteacc 
            Caption         =   "Restore and delete account"
            Visible         =   0   'False
         End
         Begin VB.Menu nmuploadtempdata 
            Caption         =   "Upload Temporary Data"
            Visible         =   0   'False
         End
         Begin VB.Menu nmg31 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu nmbackup 
            Caption         =   "Backup data tabel backup"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnsubmarkup 
         Caption         =   "Upload For Lock Account"
         Visible         =   0   'False
      End
      Begin VB.Menu mnagent 
         Caption         =   "&TeleCollection"
      End
      Begin VB.Menu mnspv 
         Caption         =   "&Supervisor"
         Visible         =   0   'False
      End
      Begin VB.Menu MnTarget 
         Caption         =   "&Target"
         Visible         =   0   'False
      End
      Begin VB.Menu MnDsr 
         Caption         =   "Daily Submission Report"
         Visible         =   0   'False
      End
      Begin VB.Menu MnWpi 
         Caption         =   "&Weekly Performance Indicator"
         Visible         =   0   'False
      End
      Begin VB.Menu MnTglSeThn 
         Caption         =   "Set Tang&gal Setahun"
         Visible         =   0   'False
      End
      Begin VB.Menu Ofl 
         Caption         =   "Open F Login"
         Visible         =   0   'False
      End
      Begin VB.Menu test1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnAccDecSub 
         Caption         =   "Accept Decline Submission"
         Visible         =   0   'False
      End
      Begin VB.Menu mnNact 
         Caption         =   "&Status Call"
         Visible         =   0   'False
      End
      Begin VB.Menu MnCCode 
         Caption         =   "Com&plaint Code"
         Visible         =   0   'False
      End
      Begin VB.Menu mndata 
         Caption         =   "&Data Quality"
         Visible         =   0   'False
      End
      Begin VB.Menu mnreason 
         Caption         =   "&Uncontacted Status Call "
         Visible         =   0   'False
      End
      Begin VB.Menu mncontacted 
         Caption         =   "&Contacted Status Call"
         Visible         =   0   'False
      End
      Begin VB.Menu mndata2 
         Caption         =   "&Campaign"
         Visible         =   0   'False
      End
      Begin VB.Menu mnProduct 
         Caption         =   "&Product"
         Visible         =   0   'False
      End
      Begin VB.Menu mnPr 
         Caption         =   "Set Product &Knowledge"
         Visible         =   0   'False
      End
      Begin VB.Menu MnBb 
         Caption         =   "Bulletin &Board"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdInfo 
         Caption         =   "Product Info"
         Visible         =   0   'False
      End
      Begin VB.Menu sbgaris 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnblack 
         Caption         =   "Black List No telpon"
         Visible         =   0   'False
      End
      Begin VB.Menu subupdate 
         Caption         =   "Update status"
         Visible         =   0   'False
      End
      Begin VB.Menu MnBlokData 
         Caption         =   "Blok Data"
         Visible         =   0   'False
         Begin VB.Menu mnblokspv 
            Caption         =   "&Schedule Blok Data"
         End
      End
      Begin VB.Menu nmSchLocktl 
         Caption         =   "Schedule Blok Data"
         Visible         =   0   'False
      End
      Begin VB.Menu setspv 
         Caption         =   "Set Target From SPV"
         Visible         =   0   'False
      End
      Begin VB.Menu mnsubahstsacc 
         Caption         =   "Ubah Status Account"
         Visible         =   0   'False
      End
      Begin VB.Menu nmformceksts 
         Caption         =   "Cek Account Status Progress"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlistreqform 
         Caption         =   "Viewer List Request Form"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlstreqnumber 
         Caption         =   "Approval Request Additional Phone"
      End
      Begin VB.Menu nmmenuformlistconfidence 
         Caption         =   "Form List Confidence"
         Visible         =   0   'False
      End
      Begin VB.Menu mnbalance 
         Caption         =   "Payment Pattern"
         Visible         =   0   'False
      End
      Begin VB.Menu separ 
         Caption         =   "-"
      End
      Begin VB.Menu mnrptsms 
         Caption         =   "Report SMS NEW"
      End
      Begin VB.Menu VSMS 
         Caption         =   "Verify SMS"
      End
      Begin VB.Menu smsblast 
         Caption         =   "Blast SMS Text"
      End
      Begin VB.Menu nmlistsmsscript 
         Caption         =   "List sms script"
      End
      Begin VB.Menu nmapprovreject 
         Caption         =   "Approved and rejected sms"
      End
      Begin VB.Menu nmblastsmsexcel 
         Caption         =   "Send SMS Blast Via Excel"
      End
      Begin VB.Menu nmg10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MNUOFFER 
         Caption         =   "Form Offering"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuploadskip 
         Caption         =   "Upload Skip Tracer"
         Visible         =   0   'False
      End
      Begin VB.Menu nmg 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu nmReportCall 
         Caption         =   "Report Call"
         Visible         =   0   'False
         Begin VB.Menu nmRptCallServer4 
            Caption         =   "Report Call Server 4"
         End
         Begin VB.Menu nmReportCallServer5 
            Caption         =   "Report Call Server 5"
         End
      End
      Begin VB.Menu nmg3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlistsendcpa 
         Caption         =   "List Send CPA"
         Visible         =   0   'False
      End
      Begin VB.Menu nmuploadcpaptp 
         Caption         =   "Upload CPA dan PTP"
         Visible         =   0   'False
      End
      Begin VB.Menu nmListUnValidNumber 
         Caption         =   "List Unvalid Number"
      End
      Begin VB.Menu nmAksesLayanaTelkom 
         Caption         =   "Akses Layanan Telkom"
      End
      Begin VB.Menu nmlistreqptp 
         Caption         =   "&List Request PTP"
      End
      Begin VB.Menu nmresetpass 
         Caption         =   "&Reset Password"
      End
      Begin VB.Menu nmReportProblemHeadset 
         Caption         =   "List Report Problem Headset"
         Visible         =   0   'False
      End
      Begin VB.Menu nmListReportProblemTelepon 
         Caption         =   "List Report Problem Telepon"
         Visible         =   0   'False
      End
      Begin VB.Menu nmblokaplikasitins 
         Caption         =   "Blok Aplikasi TINS"
      End
      Begin VB.Menu nmManageDistribusiAccount 
         Caption         =   "Manage Distribusi Account"
      End
      Begin VB.Menu mnListAccountLunas 
         Caption         =   "List Account Lunas"
      End
      Begin VB.Menu mn_list_complaint 
         Caption         =   "List Data Complaint"
      End
      Begin VB.Menu mn_list_sid 
         Caption         =   "List SID"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Distribusi Data"
      Index           =   2
      Begin VB.Menu mnup 
         Caption         =   "Distribut Otomatis"
      End
      Begin VB.Menu mnhslupload 
         Caption         =   "&Hasil Upload"
         Visible         =   0   'False
      End
      Begin VB.Menu mndist 
         Caption         =   "Hasil &Distribusi"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Pesan"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnsend 
         Caption         =   "&Kirim"
      End
      Begin VB.Menu mnbaca 
         Caption         =   "&Baca"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&About"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnProductKnowLedge 
         Caption         =   "Product &Knowledge"
         Visible         =   0   'False
      End
      Begin VB.Menu mnabout 
         Caption         =   "&Tentang Kami"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Distribusi S&TP"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu msdisstp 
         Caption         =   "Distribusi Data STP"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Distribusi &SPV"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu UPTMGM 
         Caption         =   "Distribusi Data MGM"
      End
      Begin VB.Menu UPTSTP 
         Caption         =   "Distribusi STP"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Distr&ibusi Data Tarik"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu TarikMGM 
         Caption         =   "Data CH (Player)"
      End
      Begin VB.Menu TarikLeads 
         Caption         =   "Data Leads"
      End
      Begin VB.Menu TarikStp 
         Caption         =   "Data STP"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Duplikasi"
      Index           =   8
      Visible         =   0   'False
      Begin VB.Menu MNDUPLIKASI 
         Caption         =   "Duplikasi Leads"
      End
      Begin VB.Menu MNDUPLIKASICH 
         Caption         =   "Duplikasi CH"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Pending &Duplikasi"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu MnPendingCh 
         Caption         =   "Pending CH"
      End
      Begin VB.Menu MnPendingLeads 
         Caption         =   "Pending Leads"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Kurir"
      Index           =   10
      Visible         =   0   'False
      Begin VB.Menu mnkrmaplikasi 
         Caption         =   "Kirim Aplikasi"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Report"
      Index           =   11
      Visible         =   0   'False
      Begin VB.Menu MnReportTracking 
         Caption         =   "Report Tracking"
      End
      Begin VB.Menu MnVisit 
         Caption         =   "Visit Status"
      End
      Begin VB.Menu nmReportSms 
         Caption         =   "Report SMS"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Data Confidence"
      Index           =   12
      Begin VB.Menu mn_monhly_bp 
         Caption         =   "Monthly BP (Broken Promise)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnmonthcpa 
         Caption         =   "Update Status Call"
      End
      Begin VB.Menu mnptppayment 
         Caption         =   "Update Payment"
      End
      Begin VB.Menu nmconfidenceanalisysagent 
         Caption         =   "Confidence Analisys Agent"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_confidence_list 
         Caption         =   "Confidence List"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu nmenu 
      Caption         =   "Menu"
      Visible         =   0   'False
   End
   Begin VB.Menu mntools 
      Caption         =   "&Tools"
      Begin VB.Menu mnca 
         Caption         =   "Cancel Aksesall"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu list_phone_review 
         Caption         =   "&List Phone Review"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_aoc 
         Caption         =   "&AOC"
         Visible         =   0   'False
      End
      Begin VB.Menu transfer_data 
         Caption         =   "Transfer Data"
      End
      Begin VB.Menu add_special_history 
         Caption         =   "Add Special History"
         Visible         =   0   'False
      End
      Begin VB.Menu upload_fresh_wo 
         Caption         =   "&Upload Data Fresh WO"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuploadcus 
         Caption         =   "Upload"
      End
      Begin VB.Menu mn_performance 
         Caption         =   "DeskColl Performance"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_rpt_call 
         Caption         =   "Report Call Activity"
      End
      Begin VB.Menu mn_rpt_autodial 
         Caption         =   "Report Agent Activity"
      End
      Begin VB.Menu mn_rpt_result 
         Caption         =   "Report Summary Status Call"
      End
      Begin VB.Menu mn_report_temp 
         Caption         =   "Report Temp Agent"
      End
      Begin VB.Menu mnrptloglogin 
         Caption         =   "Report Log Login && Break"
      End
      Begin VB.Menu mn_deskcoll_perform2 
         Caption         =   "Average Performance"
         Visible         =   0   'False
      End
      Begin VB.Menu mndran 
         Caption         =   "Delete & Restore Add Number"
         Visible         =   0   'False
      End
      Begin VB.Menu mndrm 
         Caption         =   "Delete & Restore Marks"
         Index           =   55
         Visible         =   0   'False
      End
      Begin VB.Menu mn_performance_reguler 
         Caption         =   "DeskColl Performance Reguler"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCallmonitor 
         Caption         =   "Call Monitor"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_copyfile 
         Caption         =   "Copy File CPA dan Dokumen Pendukung"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_option_hide 
         Caption         =   "Filter Hide System"
         Visible         =   0   'False
      End
      Begin VB.Menu mdat 
         Caption         =   "Monitoring Data Agent dan TL"
         Visible         =   0   'False
      End
      Begin VB.Menu mnpd 
         Caption         =   "Payment Deletion"
         Visible         =   0   'False
      End
      Begin VB.Menu mnappvp 
         Caption         =   "Approve Valid Phone"
      End
      Begin VB.Menu mnProductReport 
         Caption         =   "Productivity Report"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnstrategi 
         Caption         =   "Strategi"
      End
      Begin VB.Menu mnautodialer 
         Caption         =   "Setup Auto Dialer"
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnMenuRole 
         Caption         =   "Menu Role"
      End
   End
   Begin VB.Menu mn_update_db 
      Caption         =   "Update DB"
      Visible         =   0   'False
   End
   Begin VB.Menu mnkk 
      Caption         =   "Kamus Key"
      Visible         =   0   'False
   End
   Begin VB.Menu mnlist 
      Caption         =   "List"
      Begin VB.Menu mnBP 
         Caption         =   "BP"
      End
      Begin VB.Menu mnsegment 
         Caption         =   "Segment"
      End
      Begin VB.Menu mnrsms 
         Caption         =   "SMS"
      End
      Begin VB.Menu mnst 
         Caption         =   "System Training"
      End
      Begin VB.Menu mnVN 
         Caption         =   "Valid Number"
      End
   End
   Begin VB.Menu mnnote 
      Caption         =   "Note"
   End
   Begin VB.Menu mn2 
      Caption         =   "                                                                      "
      Enabled         =   0   'False
   End
   Begin VB.Menu mnShow 
      Caption         =   "<<<Show"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim m_TelpTglAwal As String
Dim m_TelpTglAkhir As String
Dim m_TelpAgent As String
Dim f_InsertTelp As Boolean
Dim M_LOGINRS As ADODB.Recordset
Public TXT As Integer
Public m_TelpUserId As String
Public m_TelpNoTelp As String
Public m_faxName As String
Public B_FAX As Boolean
Public m_targetview As Boolean
Public ParameterCTI  As String
Public M_DATA As New CLS_FRMCUST_CC_MGM
Dim COUNTER As Integer
Public Kalimat1 As String
Public PANJANG As Double
Dim satu As String
Dim dua As String
Dim tiga As String
Dim empat As String
Dim KelapKelip As Integer
Public KuotaSms As Integer
Dim STATUS As String

'@@ 13-04-2011 Tambahan buat jumlah maksimal request koneksi
Dim JmlKoneksiReq As Integer
Dim MaxKoneksiReq As Integer

'@@=== 15-12-2010 buat lock data , (di running cuma di tl)
Dim TotalTenthDetik, TotalDetik, TenthDetik, Detik, Menit, JAM As Integer
Dim jam1 As String
'@@=== 15-12-2010 buat lock data , (di running cuma di tl)


'===============@@ 6-12-2010 buat tooltip kalo ada sms yang masuk ===========================
Private Declare Function CreateWindowEx Lib "user32" _
Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
ByVal lpClassName As String, ByVal lpWindowName As _
String, ByVal dwStyle As Long, ByVal x As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight _
As Long, ByVal hWndParent As Long, ByVal hMenu _
As Long, ByVal hInstance As Long, lpParam As Any) _
As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Function GetClientRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function DestroyWindow Lib "user32" _
(ByVal hwnd As Long) As Long

'UDT (User Defined Type) RECT.
'Digunakan untuk pengaturan batas dari jendela tooltip.
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'UDT TOOLINFO.
'Digunakan untuk menentukan semua tanda yang diperlukan
'untuk membuat jendela tooltip.
Private Type TOOLINFO
  cbSize As Long
  uFlags As Long
  hwnd As Long
  uid As Long
  RECT As RECT
  hinst As Long
  lpszText As String
  lParam As Long
End Type

'Sebuah konstanta yang digunakan untuk menghubungkan
'ke fungsi API yang bernama: CreateWindowEx.
'Hal ini untuk menandakan nilai default yang digunakan.
Private Const CW_USEDEFAULT = &H80000000

'Konstanta untuk fungsi API bernama: SetWindowPosition.
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1

'Konstanta untuk menentukan gaya dari jendela tooltip.
Private Const WS_POPUP = &H80000000
Private Const WS_EX_TOPMOST = &H8&

'Konstanta yang digunakan dengan fungsi API SendMessage
'untuk mendefinisikan pesan private.
Private Const WM_USER = &H400

'Messages yang digunakan untuk menentukan durasi waktu 'dari tooltips. Tidak digunakan di sini.
Private Const TTDT_AUTOMATIC = 0
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTDT_RESHOW = 1

'Semua "penanda" untuk jendela tooltip.
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_CENTERTIP = &H2
Private Const TTF_DI_SETITEM = &H8000
Private Const TTF_IDISHWND = &H1
Private Const TTF_RTLREADING = &H4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_TRANSPARENT = &H100

'Semua pesan yang tersedia untuk tooltip Windows.
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ADDTOOLW = (WM_USER + 50)
Private Const TTM_ADJUSTRECT = (WM_USER + 31)
Private Const TTM_DELTOOLA = (WM_USER + 5)
Private Const TTM_DELTOOLW = (WM_USER + 51)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETBUBBLESIZE = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
Private Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
Private Const TTM_GETDELAYTIME = (WM_USER + 21)
Private Const TTM_GETMARGIN = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_GETTEXTW = (WM_USER + 56)
Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_GETTOOLINFOA = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW = (WM_USER + 53)
Private Const TTM_HITTESTA = (WM_USER + 10)
Private Const TTM_HITTESTW = (WM_USER + 55)
Private Const TTM_NEWTOOLRECTA = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW = (WM_USER + 52)
Private Const TTM_POP = (WM_USER + 28)
Private Const TTM_RELAYEVENT = (WM_USER + 7)
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_SETMARGIN = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLEA = (WM_USER + 32)
Private Const TTM_SETTITLEW = (WM_USER + 33)
Private Const TTM_SETTOOLINFOA = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
Private Const TTM_TRACKACTIVATE = (WM_USER + 17)
Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_UPDATE = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16)

'Konstanta untuk menentukan gaya dari jendela tooltip.
'Selalu tip, walalupun jika jendela utama tidak aktif.
Private Const TTS_ALWAYSTIP = &H1
'Menggunakan gaya balon tooltip.
Private Const TTS_BALLOON = &H40
'Win98 and up - jangan gunakan sliding tooltips.
Private Const TTS_NOANIMATE = &H10
'Win2K and up - jangan hilangkan tooltips.
Private Const TTS_NOFADE = &H20
'Mencegah Windows dari penghapusan karakter ampersand 'apapun di dalam string tooltip. Tanpa penanda ini, 'Windows otomatis akan menghapus karakter ampersand 'dari string tersebut. Hal ini dilakukan untuk 'mengizinkan string yang sama dapat digunakan
'sebagai teks dari tooltip, dan sebagai tulisan dari 'sebuah control.
Private Const TTS_NOPREFIX = &H2

'Class untuk dua tooltip yang berbeda.
Private Const TOOLTIPS_CLASS = "tooltips_class"
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'Sebuah variabel bertipe Long untuk menyimpan hwnd '(window handle) dari jendela tooltip yang dibuat di 'contoh ini.Hal ini akan menjadi sebuah array bertipe 'Long jika kita membuat tooltip Windows untuk banyak 'control atau banyak jendela.
Dim hwndTT As Long
'===============@@ 6-12-2010 buat tooltip kalo ada sms yang masuk ===========================

'@@ 13-07-2012, Buat Verifikasi form yang hanya boleh dibuka di admin
Public CekVerifikasi As Boolean



 
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000
 
Private Const WS_THICKFRAME = &H40000
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Sub add_special_history_Click()
    form_add_history.Show vbModal
End Sub

Private Sub apaja_Click()
Form3.Show
End Sub

Private Sub cmd_break_Click()
On Error GoTo klik_error
'    If MsgBox("Waktu istirahat akan diset sekarang juga??", vbYesNo + vbQuestion, "Confirm") = vbYes Then
'        M_OBJCONN.execute "UPDATE usertbl SET f_break=1"
'        MsgBox "Waktu istirahat telah diset", vbOKOnly + vbInformation, "INFO"
'    Else
'        MsgBox "Waktu istirahat dibatalkan", vbOKOnly + vbInformation, "INFO"
'    End If
    break_time = True
    frm_list_autodialer_break.Show 1
    'Timer_Durasi_status_autodialer.Enabled = False
    If AutoDialerON = False Then
        Cmdautodialer(0).Enabled = True
        Cmdautodialer(2).Enabled = False
    Else
        Cmdautodialer(0).Enabled = False
        Cmdautodialer(2).Enabled = True
    End If
klik_error:
    MsgBox err.Description
End Sub

Private Sub Cmdautodialer_Click(Index As Integer)

Select Case Index

Case 0  ' START autodialer
    Session_AutoDial = waktu_server_sekarang
    sqll = "INSERT into tbl_autodialer_agent_break(sessionid,status_break,agent,waktu_start,ip_login)values"
    sqll = sqll + " ('" + Session_AutoDial + "','AutoDial','" + MDIForm1.Text1.text + "','" & waktu_server_sekarang & "','" & Winsock1.LocalIP & "')"
    M_OBJCONN.execute sqll
    If Session_ManualDial <> "" Then
        M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end = '" & waktu_server_sekarang & "', durasi = '" & waktu_server_sekarang & "'::timestamp - '" + Session_ManualDial + "' where sessionid= '" + Session_ManualDial + "' and agent='" + MDIForm1.Text1.text + "'"
        Session_ManualDial = ""
    End If
    If break_time = False Then
        Autodialer_Start MDIForm1.Text1.text, "ManualDial", lblautdialer_timer_start_stop.Caption, ""
    Else
        Autodialer_Start "", "", "", ""
    End If
    Cmdautodialer(0).Enabled = False
    Cmdautodialer(1).Enabled = True
    Cmdautodialer(2).Enabled = True
    lblautdialer_timer_start_stop.Caption = "0"
    AutoDialerON = True
Case 1  ' Stop Autodialer
'    frm_list_autodialer_break.Show 1
'    Cmdautodialer(0).Enabled = True
'    Cmdautodialer(1).Enabled = False
Case 2  ' Manual Dialer
    Session_ManualDial = waktu_server_sekarang
    sqll = "INSERT into tbl_autodialer_agent_break(sessionid,status_break,agent,waktu_start,ip_login)values"
    sqll = sqll + " ('" + Session_ManualDial + "','ManualDial','" + MDIForm1.Text1.text + "','" & waktu_server_sekarang & "','" & Winsock1.LocalIP & "')"
    M_OBJCONN.execute sqll

     'update session autodialer
    If Session_AutoDial <> "" Then
        M_OBJCONN.execute " update tbl_autodialer_agent_break set waktu_end = '" & waktu_server_sekarang & "', durasi = '" & waktu_server_sekarang & "'::timestamp - '" + Session_AutoDial + "' where sessionid= '" + Session_AutoDial + "' and agent='" + MDIForm1.Text1.text + "'"
        Session_AutoDial = ""
    End If
    If break_time = False Then
        Autodialer_Stop MDIForm1.Text1.text, "start_autodialer", lblautdialer_timer_start_stop.Caption, "", ""
    Else
        Autodialer_Stop MDIForm1.Text1.text, "start_autodialer", lblautdialer_timer_start_stop.Caption, "", ""
    End If
    'Autodialer_Stop MDIForm1.Text1.text, "start_autodialer", lblautdialer_timer_start_stop.Caption
    Cmdautodialer(0).Enabled = True
    Cmdautodialer(1).Enabled = False
    Cmdautodialer(2).Enabled = False
    AutoDialerON = False
    lblautdialer_timer_start_stop.Caption = "0"
End Select

End Sub

Private Sub cmdenabledptp_Click()
'    Dim sql As String
'
'    If cmdenabledptp.Left = 480 Then
'       cmdenabledptp.Left = 0
'       sql = "UPDATE enabledptp SET enabled = 0"
'    Else
'       cmdenabledptp.Left = 480
'       sql = "UPDATE enabledptp SET enabled = 1"
'    End If
'    M_OBJCONN.Execute sql
'
'    Call enabledptp
End Sub

Private Sub enabledptp()
    Dim sql As String
    Dim M_objrs As ADODB.Recordset
    
    If bcp = False Then
        sql = "SELECT * FROM enabledptp"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        If M_objrs(0) = 0 Then
           'cmdenabledptp.Left = 0
           Option1.Value = 1
           Option2.Value = 0
        Else
           'cmdenabledptp.Left = 480
           Option1.Value = 0
           Option2.Value = 1
        End If
    End If
    'M_OBJCONN.Execute sql
End Sub

Private Sub Command1_Click()
    'Frm_DailyAplikasi.Show
    SSPanel_browse.Visible = False
End Sub

Private Sub CmdAddLeads_Click()
'        With FrmEntryReff
'            .TxtIdReff.Text = "Inbound Leads"
'            .TxtIdReff.Enabled = False
'             .Show vbModal
'             If .okReff Then
'             Else
'             End If
'        End With
End Sub

Private Sub CmdInbox_Click()
    FrmInboXSms.Caption = "SMS 1"
    FrmInboXSms.Show vbModal
End Sub

Private Sub Command2_Click()
        SqlWaktu = "select now()"
        Set m_waktuserver = New ADODB.Recordset
        m_waktuserver.CursorLocation = adUseClient
        m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If Val(Format(m_waktuserver(0), "hh")) >= 12 And Val(Format(m_waktuserver(0), "hh")) <= 13 Then
        SSPanel_browse.Visible = True
    Else
        MsgBox "Bacaan dapat diakses dari jam 12 S.d. 13!", vbOKOnly + vbInformation, "Informasi"
    End If
End Sub

Private Sub Command3_Click()
    FormupdateDB.Show vbModal
End Sub

Private Sub Command4_Click()
    Form4.Show
End Sub

Private Sub Command5_Click()
    query = "update dataperagent set jml = '" + Text8.text + "'"
    M_OBJCONN.execute query
    
    MsgBox "Data per agent berhasil diubah"
End Sub

Private Sub Command6_Click()
'    Strsql = " select username, extract(EPOCH from start_time) start_time, extract(EPOCH from end_time) end_time, id from ( " & vbCrLf
'    Strsql = Strsql + " SELECT * FROM public.dblink  " & vbCrLf
'    Strsql = Strsql + "      ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''2018-02-23 00:00:00'' and ''2018-02-23 23:59:59'' and username <> ''''')   " & vbCrLf
'    Strsql = Strsql + "          AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)  " & vbCrLf
'    Strsql = Strsql + "          ) a order by username, id "
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    'select username, id , start_time, end_time
'    'ulang = 1
'    userawal = rs!UserName
'    WaktuAwal = rs!end_time
'    rs.MoveNext
'    q = "Delete from tblwrapidle_temp;"
'    For i = 1 To rs.RecordCount - 1
'        If userawal = rs!UserName Then
'            q = "insert into tblwrapidle_temp values ('" & rs!UserName & "', '" & rs!start_time & "', '" & WaktuAwal & "');"
'            M_OBJCONN.Execute q
'            WaktuAwal = rs!end_time
'        Else
'            userawal = rs!UserName
'            WaktuAwal = rs!end_time
'        End If
'        rs.MoveNext
'    Next i
End Sub

Private Sub stshidenaksesall()
    On Error GoTo bawah:
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "select * from tblhidenaksesallsts"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs!STATUS = 0 Then
        Text9.text = "Nonaktif"
    Else
        Text9.text = "Aktif"
    End If
bawah:
End Sub

Private Sub Command7_Click()
    If Text9.text = "Nonaktif" Then
        qu = "Update tblhidenaksesallsts set status = 1;"
        M_OBJCONN.execute qu
        Text9.text = "Aktif"
    Else
        qu = "Update tblhidenaksesallsts set status = 0;"
        M_OBJCONN.execute qu
        Text9.text = "Nonaktif"
    End If
End Sub

Private Sub Command8_Click()
    Call stshidenaksesall
End Sub

Private Sub Label10_Click()
    If UCase(Text1) <> "ADMIN" Then
        'Load frm_showsms
        'frm_showsms.Show vbModal
        FrmInboXSms.Caption = "SMS"
        FrmInboXSms.Show vbModal
    End If
End Sub

Private Sub Label12_Click()
    FormupdateDB.Show
End Sub

Private Sub Label9_Click()
    If UCase(Text1) <> "ADMIN" Then
        'Load frm_showsms
        'frm_showsms.Show vbModal
        FrmInboXSms.Caption = "SMS"
        FrmInboXSms.Show vbModal
    End If
End Sub

Private Sub LblBersihkan_Click()
    Dim a As String
    a = InputBox("P?", "P")
    If a = "DNN#123" Then
        FrmBersihkanNegoPTP.Show vbModal
    End If
End Sub

Private Sub LblJmlSmsBaru_Change()
    Label9 = "SMS BARU " & LblJmlSmsBaru.Caption & " SMS"
End Sub

Private Sub list_phone_review_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        MsgBox "Mohon maaf, anda tidak memiliki akses!", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Then
            Form_List_Phone_Review.Show vbModal
    End If
End Sub

Private Sub LstGrade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    MDIForm1.LstGrade.SortKey = ColumnHeader.Index - 1
    MDIForm1.LstGrade.Sorted = True
End Sub

Private Sub LstGrade_DblClick()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    If LstGrade.ListItems.Count = 0 Then
    Else
        shedulePTP_Show = True
        If UCase(MDIForm1.Text2.text) = "AGENT" Then
            If UCase(MDIForm1.Text1.text) <> Trim(UCase(MDIForm1.LstGrade.SelectedItem.SubItems(3))) Then
                MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "Aplikasi"
                Exit Sub
            End If
        End If
        If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Then
        'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        Dim PO_AGENT As String
        If VIEW_MGMDATA.Combo1(0).text = "PULLOUT" Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            CMDSQL = "SELECT PO_Agent FROM mgm where CUSTID='" & MDIForm1.LstGrade.SelectedItem.SubItems(1) & "'"
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            PO_AGENT = M_objrs!PO_AGENT
            Set M_objrs = Nothing
        Else
            PO_AGENT = MDIForm1.LstGrade.SelectedItem.SubItems(3)
        End If
    
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL = "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.Text1.text + "' AND USERID = '" + PO_AGENT + "'"
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount <> 0 Then
        Else
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set M_objrs = Nothing
            Exit Sub
        End If
        Set M_objrs = Nothing
    End If
    Me.MousePointer = vbHourglass
    Flag_mgm = False
    FrmCC_Colection.Show
    Me.MousePointer = vbNormal
    'frmCC_Colection.Show
End If
End Sub

Public Sub LoOut_Ext(number$)
    Dim cancelflag As Boolean
    Dim DialString$, FromModem$, dummy
    DialString$ = "ATDT" + number$ + ";" + vbCr
    On Error Resume Next
    If MSComm1.PortOpen Then
    Else
        If MDIForm1.TxtCommPort.text = Empty Then
            MsgBox "Tidak Ada Variable buat Comport", vbInformation + vbOKOnly
            Exit Sub
        End If
        MSComm1.CommPort = MDIForm1.TxtCommPort.text
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.PortOpen = True
    End If
    Me.MousePointer = 11
    If err Then
        MsgBox err.Description, vbCritical + vbOKOnly, "Aplikasi"
        MSComm1.PortOpen = False
        cancelflag = True
        Me.MousePointer = 0
        Exit Sub
    End If
    MSComm1.InBufferCount = 0
    MSComm1.output = DialString$
    Me.MousePointer = 0
    Do
        dummy = DoEvents()
        If MSComm1.InBufferCount Then
            FromModem$ = FromModem$ + MSComm1.Input
            If InStr(FromModem$, "OK") Then
          '      Beep
                WaitSecs (0.1)
                cancelflag = True
                Exit Do
            End If
            If InStr(FromModem$, "NO DIALTONE") Then
          '      Beep
          '      Beep
                MsgBox err.Description, vbInformation + vbOKOnly, "Aplikasi"
                cancelflag = True
                Exit Do
            End If
        End If
        If cancelflag Then
            cancelflag = False
            Me.MousePointer = 0
            Exit Do
        End If
    Loop
    If MSComm1.PortOpen = True And cancelflag = True Then
        MSComm1.output = "ATH" + vbCr
        MSComm1.PortOpen = False
    End If
    Me.MousePointer = 0
End Sub

Private Sub mdat_Click()
    Formmonitoringdataagentl.Show
End Sub

Private Sub MDIForm_Activate()
    JmlKoneksiReq = 0
    MaxKoneksiReq = 200
        
    Call enabledptp
    Call stshidenaksesall
    ' set button autdialer manual dahulu
    
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        cmd_break.Visible = True
    End If
    
    If UCase(Trim(MDIForm1.Text2.text)) <> "SUPERVISOR" And UCase(Trim(MDIForm1.Text2.text)) <> "MANAGER" Then
        'Disable Close Button
        'DisableCloseBtn Me
        '=======================
        Dim hMenu   As Long
        Dim lStyle  As Long
    
        'disable MAXIMIZE button
        lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
        lStyle = lStyle And Not WS_MINIMIZE
        lStyle = lStyle And Not WS_MAXIMIZEBOX
        Call SetWindowLong(Me.hwnd, GWL_STYLE, lStyle)
    End If
    
    DisableCloseBtn Me
    'jejaktian12042016
    'SSPanel5.Width = 100
    '================================
   'FrmConfidenceAnalysis.Show vbModal
   
    STATUS = "Connected"
    'Call insertlogcti(STATUS)
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strsql As String
    strsql = "UPDATE usertbl SET stsaplikasi=0 WHERE userid ='" + MDIForm1.Text1.text + "'"
    M_OBJCONN.execute (strsql)
    
    '@@ 13-04-2011 Hapus data ip
    strsql = "delete from tbl_ip where agent='"
    strsql = strsql + Trim(MDIForm1.Text1.text) + "'"
    M_OBJCONN.execute strsql
    
    Winsock2.Close
    
    '@@28012013 ini buat update status loginnya
    strsql = "UPDATE usertbl SET f_status_login=null,last_logout='now()' WHERE userid='"
    strsql = strsql + Trim(MDIForm1.Text1.text) + "'"
    M_OBJCONN.execute strsql
    
    Call set_count_ol("log out")
    
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    strsql = "UPDATE usertbl SET  f_status_login=null,last_logout='now()' WHERE userid='"
    strsql = strsql + Trim(MDIForm1.Text1.text) + "'"
    M_OBJCONN.execute strsql
     
    Cancel = 1
    
    
End Sub

Private Sub mn_aoc_Click()
    FormAOC.Show vbModal
End Sub

Private Sub mn_confidence_list_Click()
    FrmConfidenceList.Show 1
End Sub

Private Sub mn_copyfile_Click()
    Form_CopyFIleCPA.Show vbModal
End Sub

Private Sub mn_deskcoll_perform2_Click()
    Form_deskcoll_performance2.Show 1
End Sub

Private Sub mn_list_complaint_Click()
    Frm_list_complaint.Show 1
End Sub

Private Sub mn_list_sid_Click()
    CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
        UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
        UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Or _
        UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        
        FrmVerifikasiPassword.txtUserName.text = UCase(MDIForm1.Text1.text)
        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            FrmSID.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
End Sub

Private Sub mn_monhly_bp_Click()
    Form_monthly_BP.Show 1
End Sub

Private Sub mn_option_hide_Click()
    Form_filter_hide.Show 1
End Sub

Private Sub mn_performance_Click()
    Form_deskcoll_performance.Show 1
End Sub

Private Sub mn_performance_reguler_Click()
    Form_deskcoll_performance_reguler.Show 1
End Sub

Private Sub mn_report_call_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frm_report_call.Show vbModal
End Sub

Private Sub mn_report_temp_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
        Form_report_temp.Show vbModal
End Sub

Private Sub mn_rpt_autodial_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
        Form_Report_AutoDial.Show vbModal
End Sub

Private Sub mn_rpt_call_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frm_report_call.Show 1
End Sub

Private Sub mn_rpt_result_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
        Form_report_summary.Show 1
End Sub
Private Sub mn_update_db_Click()
    On Error Resume Next
    M_OBJCONN.execute "CREATE TABLE tbl_count_block(" & _
                        "tgl timestamp without time zone default now()," & _
                        "agent varchar(30)," & _
                        "ket VarChar(150));"
                            
    MsgBox "Update Database Successfully !!!"
End Sub

Private Sub mnappvp_Click()
    form_approvevalid.Show
End Sub

Private Sub mnautodialer_Click()

    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or _
       Left(MDIForm1.Text2.text, 2) = "AM" Then
        'FrmVerifikasiPassword.TxtUsername.Text = UCase(MDIForm1.Text1.Text)
        'FrmVerifikasiPassword.Show vbModal
        
'        If CekVerifikasi = True Then
            frm_list_autodialer_setup.Show vbModal
'        Else
'            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
'        End If
    End If

End Sub

Private Sub mnbalance_Click()
    FrmBalance.Show vbModal
End Sub

Private Sub MnBb_Click()
'    FRM_Bulletin_LIST.Show vbModal
End Sub

Private Sub mnblack_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
        FrmVerifikasiPassword.txtUserName.text = UCase(MDIForm1.Text1.text)
        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            frm_BlackListNo_List.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
    
    'frm_BlackListNo_List.Show 1
End Sub

Private Sub mnblokspv_Click()
    'frmlockaccountfromspv.Show 1
    frm_map_lock_acc.Show 1
End Sub

Private Sub mnBP_Click()
    formlistbp.Show 'vbModal
End Sub

Private Sub mnca_Click()
    frmreleaseaksesall.Show vbModal
End Sub

Private Sub MnCCode_Click()
    FRM_Complaint_LIST.Show vbModal
End Sub

Private Sub mncontacted_Click()
    FRM_ContactedDesc_LIST.Show vbModal
End Sub

Private Sub mndata_Click()
  '  FRM_DataQuality_LIST.Show vbModal
End Sub


Private Sub mndran_Click()
    Formdelresaddnumber.Show
End Sub

Private Sub mndrm_Click(Index As Integer)
    FrmRestoreRemarks.Show vbModal
End Sub

Private Sub MnDuplikasi_Click()
   ' FrmDuplikasi.Show
End Sub

Private Sub MNDUPLIKASICH_Click()
   ' FrmDuplikasiCh.Show
End Sub

Private Sub mnkk_Click()
    Kamuskey.Show
End Sub

Private Sub mnkrmaplikasi_Click()
   ' FrmKurirSpvApp.Show vbModal
End Sub

Private Sub mnListAccountLunas_Click()
       CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        FrmVerifikasiPassword.txtUserName.text = UCase(MDIForm1.Text1.text)
        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            FrmAccLunas.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
            Exit Sub
        End If
    End If
    
    FrmAccLunas.Show vbModal
End Sub

Private Sub mnMenuRole_Click()
    Form_Menu_Role.Show vbModal
End Sub

Private Sub mnmonthcpa_Click()
    'Form_List_CPA.Show vbModal
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        MsgBox "Mohon maaf, anda tidak memiliki akses!", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Then
            frm_upstatus_call.Show vbModal
    End If
    
End Sub

Private Sub mnNact_Click()
    FRM_NextAct_LIST.Show vbModal
End Sub

Private Sub mnnote_Click()
    formnote.Show
End Sub
Private Sub mndh_Click()
    form_mndd.Show
    'FormReport_DonotCall.Show
End Sub

Private Sub mnpd_Click()
    Frmdeltbllunas.Show 1
End Sub

Private Sub mnProductReport_Click()
    frmreportproductivity.Show
End Sub

Private Sub mnptppayment_Click()
    'Form_ptp_payment.Show vbModal
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        MsgBox "Mohon maaf, anda tidak memiliki akses!", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Then
            form_tarikdata.Show vbModal
    End If
    
End Sub

Private Sub MnReportTracking_Click()
 FrmMgmReport.Show
End Sub

Private Sub mnrptloglogin_Click()
    Form_Report_LoginBreak.Show 1
    Form_Report_LoginBreak.ZOrder vbBringToFront
End Sub

Private Sub mnrptsms_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frm_report_sms1.Show 1
End Sub

Private Sub mnrsms_Click()
    If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Then
        formreportsms.Show
    End If
End Sub

Private Sub mnsegment_Click()
'    If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Then
        formsegment.Show
'    End If
End Sub

Private Sub mnShow_Click()
    If mnShow.Caption = "<<<Show" Then
        mnShow.Caption = "Hide>>>"
        SSPanel5.Visible = True
    Else
        mnShow.Caption = "<<<Show"
        SSPanel5.Visible = False
    End If

End Sub

Private Sub mnST_Click()
    If UCase(Text2.text) = "SUPERVISOR" Then
        formsystemtraining.Show 1
    End If
End Sub

Private Sub mnstrategi_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Then
        form_strategi.Show 1
    End If
End Sub

Private Sub mnsubahstsacc_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frnubahstsaccount.Show 1
End Sub

Private Sub mnsubmarkup_Click()
'    If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Then
'        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
'        Exit Sub
'    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or Left(MDIForm1.Text2.text, 2) = "AM" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
        FrmVerifikasiPassword.txtUserName.text = UCase(MDIForm1.Text1.text)
        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            FRMMARKUP.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
    
    
    'FRMMARKUP.Show 1
End Sub

Private Sub mntest_Click()

End Sub

Private Sub mnuCallmonitor_Click()
    Form_call_mon.Show
End Sub

Private Sub MNUOFFER_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmheaderoffeer.Show 1
End Sub

Private Sub mnuploadcus_Click()
    Form_upload.Show
End Sub

Private Sub mnuploadskip_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmuploadskiptarcer.Show 1
End Sub
Private Sub mnuProdInfo_Click()
    FrmProductList.Show
End Sub

Private Sub mnuuploadcpa_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    Form_upload_CPA.Show vbModal
End Sub

Private Sub MnVisit_Click()
    FrmVisit.Show
End Sub

Private Sub nmAksesLayanaTelkom_Click()
       If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmAksesLayananTelkom.Show vbModal
End Sub

Private Sub nmapprovreject_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    'frm_approved_rejected.Show vbModal
    Frm_verify.Show vbModal
End Sub

Private Sub nmbackup_Click()
     If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmBackupDbToExcel.Show vbModal
End Sub

Private Sub nmblastsmsexcel_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmSmsBlastExcel.Show
End Sub

Private Sub nmblokaplikasitins_Click()
    
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        MsgBox "Mohon maaf, anda tidak memiliki akses!", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or Left(MDIForm1.Text2.text, 2) = "AM" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Then
       'UCase(Trim(MDIForm1.Text2.Text)) = "ADMIN"
        ' REQUEST JOKO TGL. 30 SEP 2013 - Tanpa Verifikasi
        'FrmVerifikasiPassword.TxtUsername.Text = UCase(MDIForm1.Text1.Text)
        'FrmVerifikasiPassword.Show vbModal
        
'        If CekVerifikasi = True Then
            FrmBlokAgent.Show vbModal
'        Else
'            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
'        End If
    End If
End Sub

Private Sub nmconfidenceanalisysagent_Click()
    FrmConfidenceListNew_Agent.Show vbModal
End Sub

Private Sub nmenu_Click()
    FrmListReqTlp.Show vbModal
End Sub

Private Sub nmformceksts_Click()
    Frm_Cek_status_acc.Show 1
End Sub

Private Sub nmListReportProblemTelepon_Click()
    Dim a As String
    a = InputBox("Password?", "@@@@@")
    If a = "Dnn#12345" Then
        FrmListReportTelepon.Show vbModal
    Else
        MsgBox "Akses di tolak!", vbOKOnly + vbExclamation, "Peringatan"
    End If
End Sub

Private Sub nmlistreqform_Click()
    FrmListRequest.Show vbModal
End Sub

Private Sub nmlistreqptp_Click()
    FrmListRequestPTP.Show vbModal
End Sub

Private Sub nmlistsendcpa_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmSendCPA.Show vbModal
End Sub

Private Sub nmlistsmsscript_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Frm_List_SMS_Script.Show vbModal
End Sub

Private Sub nmListUnValidNumber_Click()
    FrmListUnValidNumber.Show vbModal
    
End Sub

Private Sub nmlstreqnumber_Click()
    FrmListReqTlp.Show vbModal
End Sub

Private Sub nmManageDistribusiAccount_Click()
    CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then ' Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or Left(MDIForm1.Text2.text, 2) = "AM" Or _
    UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
           UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Then
        'FrmVerifikasiPassword.TxtUsername.Text = UCase(MDIForm1.Text1.Text)
        'FrmVerifikasiPassword.Show vbModal
        
        'If CekVerifikasi = True Then
            FrmDistribusiAcc.Show
        'Else
        '    MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        'End If
    End If
    
End Sub

Private Sub nmmenuformlistconfidence_Click()
    FrmConfidenceListNew.Show vbModal
End Sub

Private Sub nmreportcall_Click()
    'rptCallTracking.Show vbModal
End Sub

Private Sub nmReportCallServer5_Click()
    RptCallTrackingServer5.Show vbModal
End Sub

Private Sub nmReportProblemHeadset_Click()
    Dim a As String
    a = InputBox("Password?", "@@@@@")
    If a = "Dnn#12345" Then
        FrmListProblemHeadset.Show vbModal
    Else
        MsgBox "Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
    End If
End Sub

Private Sub nmReportSms_Click()
    frm_report_sms.Show vbModal
End Sub

Private Sub nmresetpass_Click()
    FrmResetPass.Show vbModal
End Sub

Private Sub nmrestoredeleteacc_Click()
     If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmRestoreDelete.Show vbModal
End Sub

Private Sub nmRptCallServer4_Click()
    rptCallTrackingServer4.Show vbModal
End Sub

Private Sub nmSchLocktl_Click()
'DIBUKA LAGI BY REQUEST DODDY 5-6-2015
'    If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Then
'        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
'        Exit Sub
'    End If
'
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or _
       Left(MDIForm1.Text2.text, 2) = "AM" Then
        'FrmVerifikasiPassword.TxtUsername.Text = UCase(MDIForm1.Text1.Text)
        'FrmVerifikasiPassword.Show vbModal
        
'        If CekVerifikasi = True Then
            frm_list_schedule_tl.Show vbModal
'        Else
'            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
'        End If
    End If
    
    'frm_list_schedule_tl.Show 1
End Sub



Private Sub nmswapdata_Click()
    CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
    UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
           UCase(Trim(MDIForm1.Text2.text)) = "MANAGER" Then
        FrmVerifikasiPassword.txtUserName.text = UCase(MDIForm1.Text1.text)
        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            Form_swap.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
End Sub

Private Sub nmuploadcpaptp_Click()
    FrmUploadCPAPTP.Show vbModal
End Sub

Private Sub nmuploadcustomer_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Form_upload.Show vbModal
End Sub

Private Sub nmuploadpayment_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Form_upload_payment.Show vbModal
End Sub

Private Sub nmuploadtempdata_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmUploadTempData.Show vbModal
End Sub

Private Sub Ofl_Click()
    form_open_login.Show
End Sub

Private Sub Option1_Click()
    Option2.Value = 0
    sql = "UPDATE enabledptp SET enabled = 0"
    M_OBJCONN.execute sql
End Sub

Private Sub Option2_Click()
    Option1.Value = 0
    sql = "UPDATE enabledptp SET enabled = 1"
    M_OBJCONN.execute sql
End Sub

Private Sub setspv_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmsettarget.Show 1
End Sub

Private Sub smsblast_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Load frm_sms_blast
    DoEvents
    frm_sms_blast.Show vbModal
    DoEvents
End Sub

Private Sub SSCommand1_Click(Index As Integer)
Dim M_objrs As ADODB.Recordset
Dim CMDSQL As String
Dim m_objrscekmonitoring As ADODB.Recordset

Select Case Index
    Case 0
        m_targetview = True
        If MDIForm1.Text2.text = "Agent" Then
            VIEW_MGMDATA.LblTarget(0).Caption = LblTarget.Caption
            VIEW_MGMDATA.LblTarget(1).Caption = LblTarget.Caption
            VIEW_MGMDATA.LstVwSearchMgm.Checkboxes = False
        End If
        
        Dim ds As New ADODB.Recordset
        
        ds.CursorLocation = adUseClient
        ds.Open "select lockdarispv, F_LOCK, f_akses_all_acc FrOM usertbl WHERE USERID='" & MDIForm1.Text1.text & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If ds.BOF And ds.EOF Then
        Else
            If ds!F_LOCK = "Y" Then
                F_LOCK = True
            Else
                F_LOCK = False
            End If
        End If
        
        If ds.EOF = False Then
            cek_aksesall = cnull(ds("f_akses_all_acc"))
        Else
            cek_aksesall = 0
        End If
        
        If cek_aksesall = "1" Then
            VIEW_MGMDATA.CmdSearchPTP.Enabled = False
        End If
        
        If cek_aksesall = "0" Then
            VIEW_MGMDATA.CmdSearchPTP.Enabled = True
        End If
        
        ' CEK LOCK ACC
        If ds.EOF = False Then
            If ds("lockdarispv") <> "" Then
                VIEW_MGMDATA.CmdSearchPTP.Enabled = False
            End If
        End If
                    
        If UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Or UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
            '@@ 10-03-2011 Tambahan buat nambahin monitoring headset
            'cek dulu apakah status monitoring aktif
            CMDSQL = "select * from manajemen_site  where status='1'"
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                If M_objrs.RecordCount > 0 Then
                    CMDSQL = "select monitoring_headset from usertbl where userid='"
                    CMDSQL = CMDSQL + Trim(MDIForm1.Text1.text) + "'"
                    Set m_objrscekmonitoring = New ADODB.Recordset
                    m_objrscekmonitoring.CursorLocation = adUseClient
                    m_objrscekmonitoring.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    If Trim(m_objrscekmonitoring("monitoring_headset")) = "1" Then
                        FrmMonitoringHeadset.Show vbModal
                        Set m_objrscekmonitoring = Nothing
                    Else
                        VIEW_MGMDATA.Show
                        Set M_objrs = Nothing
                        Set m_objrscekmonitoring = Nothing
                    End If
                Else
                    VIEW_MGMDATA.Show
                    Set M_objrs = Nothing
                End If
        Else
            VIEW_MGMDATA.Show
        End If
        'FRM_SEARCH.Show
    Case 1
        'FrmmgmReportKeDua.Show
   ' Case 2
    '    FrmPreembosReport.Show
    Case 10
        FRMSENDMSG.Show vbModal
        'FRMSENDMSG.Show
    Case 4
        
    Case 5
        FrmVisit.Show
    Case 6
'        FrmSearching.Show
        FrmCari.Show
    Case 7
        fmunlock.Show
    Case 8
        FrmAccessData.Show
    Case 9
       ' FrmmgmReport.Show
    Case 11
    FrmMgmReport.Show ' UTK RITCARD
    'FrmMgmReport_AWARNESS.Show ' utk RITPIL dan AwarNESS
End Select
End Sub

Private Sub MDIForm_Load()
    custid_autodial_not_in = ""
    Call calldataperagent
    
    '------------ klik button manual dialer ----------------
    ' set manual auotdialer dahulu
        '------------------------------
    '--------- end klik button manual dialer -------------------
    
    
    
    'NGAMBIL WAKTU LOGIN UNTUK BLOCK
    waktu_start = waktu_server_sekarang
       
      '@@ 14-02-2012, media
    '  If Val(Format(Now(), "hh")) >= 12 And Val(Format(Now(), "hh")) <= 13 Then
    '        SSPanel_browse.Visible = True
    '  End If
    
    '        WebBrowser1.Navigate ("http://localhost/sobatmuslim/lokomedia/mobile/index.php")
    
    
    COUNTER = 0
    count_timer_detik = 0
    ' MONITORING ACTIVITY BY IZUDDIN 16 04 2013
    
    i_monitoring_activity = 0
    lbl_timer_activity = 0
    open_sms = False
    ' #########################################
    
    
    Timer2.Enabled = False
    
    Dim M_DATA As New CLS_LOGIN
    Dim M_LOGINRS As ADODB.Recordset
    Dim m_port As String
    'On Error GoTo MDIForm_LoadErr
                'Winsock2.Listen
                On Error GoTo HELL
                Winsock2.Listen
                
    Call PromiseToPay
    Call HeaderInformation
    bRenderrecord = False
    Call LstDataInformation
    
    
    SSTab1.TabVisible(1) = False
    'Call tglhost
    Set M_LOGINRS = New ADODB.Recordset
    M_LOGINRS.CursorLocation = adUseClient
    M_LOGINRS.Open "SELECT * FROM vwcallcfg1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_LOGINRS.RecordCount <> 0 Then
        Text6.text = IIf(IsNull(M_LOGINRS("DELAY_TONE")), "0", M_LOGINRS("DELAY_TONE"))
        TxtAuthPrefix.text = IIf(IsNull(M_LOGINRS("AUTHPREFIX")), "", M_LOGINRS("AUTHPREFIX"))
        TxtModemAcod.text = IIf(IsNull(M_LOGINRS("MODEMACOD")), "", M_LOGINRS("MODEMACOD"))
        TxtCommPort.text = IIf(IsNull(M_LOGINRS("COMMPORT")), "", M_LOGINRS("COMMPORT"))
        TDBDate1.Value = IIf(IsNull(M_LOGINRS("TglSystem")), "", M_LOGINRS("TglSystem"))
        TxtJamMulaiTelp.text = IIf(IsNull(M_LOGINRS("JAMMULAITELP")), "", M_LOGINRS("JAMMULAITELP"))
        TxtJamSelesaiTelp.text = IIf(IsNull(M_LOGINRS("JAMSELESAITELP")), "", M_LOGINRS("JAMSELESAITELP"))
        TxtLamaFollowup.text = IIf(IsNull(M_LOGINRS("LAMAFOLLOWUP")), "99", M_LOGINRS("LAMAFOLLOWUP"))
    Else
        TDBDate1.Value = Now
        TxtLamaFollowup.text = "99"
    End If
    m_port = BUKA_FILE_KONEKSI("comport.txt")
    If m_port <> "" Then
        TxtCommPort.text = m_port
    End If
    M_LOGINRS.Close
    Set M_LOGINRS = Nothing
'----------------------------------------------------------------------------------------
'M_DATA.UPDATE_CLIENT M_OBJCONN, Winsock1.LocalIP, MDIForm1.Text1.Text, Winsock1.LocalHostName
'M_DATA.INSERT_LOGIN M_OBJCONN, TDBDate1.Text & "-" & CStr(Time), Winsock1.LocalHostName, "LOGIN", MDIForm1.Text1.Text
'Exit Sub
'MDIForm_LoadErr:
'    MsgBox Err.Description
'    Set M_LOGINRS = Nothing

   'Cmdautodialer_Click (2)

    Exit Sub
HELL:
    MsgBox err.Description + "only one aplication is open", vbCritical + vbOKOnly, "Warning"
    End
End Sub

Private Sub messageapptransfer()
    cm_connect.Enabled = False
    cm_disconnect.Enabled = False
    cm_send.Enabled = False
End Sub

Private Sub mnabout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnagent_Click()
    FRM_AGENT_LIST.Show vbModal
End Sub

Private Sub mnbaca_Click()
    FRMBACAMSG.Show vbModal
End Sub

Private Sub mndata2_Click()
    FRM_DATASOURCE_LIST.Show vbModal
End Sub

Private Sub MnFile_Click(Index As Integer)
Dim strsql As String
Select Case Index
    Case 0
        Unload MDIForm1
        frmlogin.Show vbModal
    Case 1
        FRM_SET_PWD.Show
    Case 3
    Case 5
        frm_gantipas.Text1(2).text = UCase(MDIForm1.Text1.text)
        frm_gantipas.Show vbModal
    Case 6
    Case 7
    
    Call offsesilogin(MDIForm1.Text1.text)
    Call offsesilogin_new(MDIForm1.Text1.text)
    
    '@@ 13-04-2011 Hapus data ip
        strsql = "delete from tbl_ip where agent='"
        strsql = strsql + Trim(MDIForm1.Text1.text) + "'"
        M_OBJCONN.execute strsql
    
    Unload Me
    End
End Select
End Sub

Private Sub mnhslupload_Click()
    FRM_HASILUPLOAD.Show vbModal
End Sub

Private Sub mnproduct_Click()
    FRM_PRODUCT_LIST.Show vbModal
End Sub

Private Sub mnreason_Click()
    FRM_CLOSSING_LIST.Show vbModal
End Sub
Private Sub mnsend_Click()
    FRMSENDMSG.Show vbModal
End Sub

Private Sub mnspv_Click()
    FRM_SPV_LIST.Show vbModal
End Sub

Private Sub mnup_Click()
    'FRM_SETUSER.Show vbModal
    'Form_Distribute_Otomatis.Show vbModal
End Sub

Private Sub SSCommand2_Click()
    Dim M_objrs As New ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "SELECT custid,f_cek,agent FROM mgm where f_pending='pending' "
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If M_objrs.RecordCount <> 0 Then
        LstGrade.ListItems.clear
    End If
    While Not M_objrs.EOF
        Set ListItem = LstGrade.ListItems.ADD(, , M_objrs.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("F_CEK")), "", M_objrs("F_CEK"))
        ListItem.SubItems(3) = IIf(IsNull(M_objrs!agent), "", M_objrs!agent)
        M_objrs.MoveNext
    Wend
    If M_objrs.RecordCount = 0 Then
    LstGrade.ListItems.clear
    End If
    Set M_objrs = Nothing
End Sub

Private Sub SSCommand3_Click()
    Form_manual_dial.Show
End Sub

Private Sub SSPanel5_Click()
'    If SSPanel5.Width = 1125 Then
'        SSPanel5.Width = 100
'    Else
'        SSPanel5.Width = 1125
'    End If
End Sub

Private Sub subupdate_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmupdate.Show 1
End Sub

Private Sub calldataperagent()
atas:
If bcp = False Then
    CMDSQL = "select jml from dataperagent"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        query = "INSERT INTO dataperagent values (402);"
        M_OBJCONN.execute query
        GoTo atas:
    End If
    Text8.text = M_objrs!JML
End If
End Sub

Private Sub Timer_Durasi_status_autodialer_Timer()
lblautdialer_timer_start_stop.Caption = Val(lblautdialer_timer_start_stop.Caption) + 1
End Sub

Private Sub Timer10_Timer()
    If MDIForm1.Text2.text <> "Agent" And MDIForm1.Text2.text <> "TeamLeader" Then
        q = "SELECT * FROM enabledptp"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
        If r.RecordCount <> 0 Then
            If r!Enabled = 0 Then
                Option1.Value = True
                Option2.Value = False
            Else
                Option1.Value = False
                Option2.Value = True
            End If
        End If
    End If
End Sub

Private Sub Timer11_Timer()
    On Error GoTo bawah '20190725
    If UCase(MDIForm1.Text2.text) <> "SUPERVISOR" And UCase(MDIForm1.Text2.text) <> "ADMIN" And UCase(MDIForm1.Text2.text) <> "MANAGER" Then
        qsel = "select * from ("
        qsel = qsel & " select nama_file as training, jam_awal, jam_akhir, agent, f_done, b.ids  from tblsystemtraining_schedule a inner join"
        qsel = qsel & " tblsystemtraining_partisipan b on a.ids = b.ids inner join"
        qsel = qsel & " tblsystemtraining c on a.idp = c.id"
        qsel = qsel & " ) a where jam_awal < now() and jam_akhir > now() and agent = '" & MDIForm1.Text1.text & "' and coalesce(f_done,0) <> 1 order by jam_awal, agent limit 1"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open qsel, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If rs.RecordCount > 0 Then
            If UCase(MDIForm1.Text2.text) <> "TEAMLEADER" Then
                If signtimer2 = True Then
                    Timer2.Enabled = False
                    signtimes = "time2"
                ElseIf signtimer7 = True Then
                    Timer7.Enabled = False
                    signtimes = "time7"
                End If
                Timer11.Enabled = False
                'Timer7.Enabled = False
            End If
            formsystemtrainingagents.Show 1
        End If
    End If
bawah:
End Sub

Private Sub Timer12_Timer()
'    VIEW_MGMDATA.readstrategi_open
'    If open_strategi = True Then
'        VIEW_MGMDATA.readstrategi
'    End If
End Sub

Private Sub Timer13_Timer()
If bcp = False Then
        If MDIForm1.Text2.text <> "Agent" And MDIForm1.Text2.text <> "TeamLeader" Then
            Timer9.Enabled = True
            q = "SELECT * FROM tblvalidtospv"
            Set r = New ADODB.Recordset
            r.CursorLocation = adUseClient
            r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
            mntools.Caption = "&Tools "
            mnappvp.Caption = "Approve Valid Phone "
            If r.RecordCount <> 0 Then
                mntools.Caption = mntools.Caption & "(" & r.RecordCount & ")"
                mnappvp.Caption = mnappvp.Caption & "(" & r.RecordCount & ")"
            End If
        Else
            Timer9.Enabled = False
        End If
    End If
End Sub

Private Sub Timer6_Timer()
    'Ini buat ngecek sms apakah ada sms baru yang masuk
    '@@ 14022011
    Dim CMDSQL As String
    Dim M_objrs  As ADODB.Recordset
    
    If UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        'Cek dulu di usertbl apakah ada status sms masuk
        '@@ 14/02/2010,, Cek smsnya melalui field blink di usertbl aja, jadinya lebih ringan
        If UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
            CMDSQL = "select status_sms from usertbl where userid='"
            CMDSQL = CMDSQL + Trim(MDIForm1.Text1.text) + "'"
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
            If M_objrs("status_sms") <> "" Then
                'Ini jika ada sms masuk
                Call CekSms
            Else
                'Ini jika tidak ada sms masuk
                TimerBlink.Enabled = False
                Label9.ForeColor = vbBlack
            End If
            Set M_objrs = Nothing
            
            ' ++++++++++++++ CEK REMINDER 10 Mei 2013 IZUDDIN +++++++++++++++
            Dim str_detik As String
            Dim str_group_time As Integer
            Dim LocTextFile, CallerIDfromTextFile As String
            Dim str_time, isi As String
            Dim arr_reminder() As String
            Dim reminder_custid As String
            Dim reminder_jam As String
            Dim reminder_custname As String
            
            On Error GoTo err
            
            LocTextFile = "C:\reminder.txt"
        
            SqlWaktu = "select now()"
            Set m_waktuserver = New ADODB.Recordset
            m_waktuserver.CursorLocation = adUseClient
            m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            str_time = Format(m_waktuserver(0), "HH:MM")
            
            Set m_waktuserver = Nothing
            
            Open LocTextFile For Input As #1    'Buka file text
            Do Until EOF(1)
                Line Input #1, CallerIDfromTextFile      'Baca Baris Pertama
                isi = Replace(CallerIDfromTextFile, """", "")
                arr_reminder = Split(isi, "|")
                reminder_custid = arr_reminder(0)
                reminder_custname = arr_reminder(1)
                reminder_jam = arr_reminder(2)
                If reminder_jam = str_time Then
                    With frm_reminder
                        .Label1(4).Caption = reminder_custid
                        .Label1(5).Caption = reminder_custname
                        .Label1(6).Caption = Format(Now, "DD/MM/YYYY") & " - " & reminder_jam
                        .ZOrder 0
                        .Show vbModal
                    End With
                End If
            Loop
err:
            Close #1 'Tutup File file text
            ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Dim ConMSG As New ADODB.Connection
    Dim cmdsqlnew As String
    Dim cmdsql3 As String
    Dim M_objrs As New ADODB.Recordset
    Dim CMDSQL As String
    
    'On Error GoTo SALAH
    
    ConMSG.Open CMDSQLOPEN
    M_objrs.CursorLocation = adUseClient
    cmdsql3 = "select sender, recipient, datetime, msg, t from msgtbl where recipient ='" + Trim(MDIForm1.Text1.text) + "' and sts ='0'"
    M_objrs.Open cmdsql3, ConMSG, adOpenDynamic, adLockOptimistic, adCmdText



    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If

    While Not M_objrs.EOF
        FRMTERIMAMSG.RichTextBox1.SelColor = &HC00000
        FRMTERIMAMSG.Text1.text = IIf(IsNull(M_objrs!Sender), "", M_objrs!Sender)
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Dari :" + IIf(IsNull(M_objrs!Sender), "", M_objrs!Sender) + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Kepada :" + IIf(IsNull(M_objrs!RECIPIENT), "", M_objrs!RECIPIENT) + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Tanggal :" + IIf(IsNull(M_objrs!DateTime), "", M_objrs!DateTime) + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Isi Pesan :" + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + IIf(IsNull(M_objrs!msg), "", M_objrs!msg)
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + " " + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text & vbCrLf
        M_objrs.MoveNext
    Wend
    If M_objrs.RecordCount <> 0 Then
        'On Error GoTo SALAH
        'FRMTERIMAMSG.Show vbModal
        FRMTERIMAMSG.Show vbModal
        cmdsql3 = "UPDATE msgtbl SET STS ='1' WHERE RECIPIENT ='" + MDIForm1.Text1.text + "'"
        ConMSG.execute cmdsql3
    End If
    Set M_objrs = Nothing
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    cmdsqlnew = "select * from usertbl where userid='" + MDIForm1.Text1.text + "' and f_flagrender=1"
    M_objrs.Open cmdsqlnew, ConMSG, adOpenDynamic, adLockOptimistic

    If M_objrs.RecordCount <> 0 Then
        bRenderrecord = True
        cmdsql3 = "UPDATE usertbl SET f_flagrender =0 where userid='" + MDIForm1.Text1.text + "'"
        ConMSG.execute cmdsql3
    End If

    ConMSG.Close
    Set ConMSG = Nothing
    Exit Sub
'SALAH:
'bikin error
    'FRMTERIMAMSG.Hide
    'MsgBox "Ada error : " & err.Description

    
'    'jejaktian===========================================================================
'    ConMSG.Open CMDSQLOPEN
'    M_Objrs.CursorLocation = adUseClient
'    cmdsql3 = "select pemohon from tblpermohonantransferdata where penggaprove = ''"
'    M_Objrs.Open cmdsql3, ConMSG, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_Objrs.RecordCount = 0 Then
'        Set M_Objrs = Nothing
'        Exit Sub
'    End If
'
'    While Not M_Objrs.EOF
'        Formapptransferdata.lblpemohon.Caption = IIf(IsNull(M_Objrs!pemohon), "", M_Objrs!pemohon)
'        M_Objrs.MoveNext
'    Wend
'    If M_Objrs.RecordCount <> 0 Then
'        On Error GoTo SALAH
'        'FRMTERIMAMSG.Show vbModal
'        Formapptransferdata.Show vbModal
'        cmdsql3 = "UPDATE tblpermohonantransferdata SET penggaprove = '" + MDIForm1.Text1.Text + "' WHERE pemohon is null"
'        ConMSG.Execute cmdsql3
'    End If
'
'    ConMSG.Close
'    Set ConMSG = Nothing
'    Exit Sub
'SALAH:
'    Formapptransferdata.Hide
'    MsgBox "Ada error : " & err.Description
End Sub

Private Sub Timer2_Timer()
    signtimer2 = True
        
    Dim dura_blok As Integer

    If UCase(MDIForm1.Text2.text) <> "AGENT" Then
       i_monitoring_activity = 0
       Timer2.Enabled = False
       Exit Sub
    End If
    
'    TxtStatus.Text = "FEEDBACKhangup"
'
    If TxtStatus.text Like "*FEEDBACKhangup*" Then
        i_monitoring_activity = 0
        TxtStatus.text = ""
        Call logwktcti("RESET WAKTU BY TIAN")
    End If
    
    'If Not TxtStatus.Text Like "*FEEDBACKhangup*" Then
        'waktu_selesai_ngitung = waktu_server_sekarang
    'End If
    'TxtStatus.Text = "FEEDBACKbusy"
    
    'Call WskCTI_DataArrival(FEEDBACKbusy)
    
'    If TxtStatus.Text = "FEEDBACKhangup" Then
'        FrmCC_Colection.Label12.Caption = 0
'        TxtStatus.Text = ""
'    End If
        
    'dura_blok = DateDiff("s", waktu_mulai_ngitung, waktu_selesai_ngitung)

    i_monitoring_activity = i_monitoring_activity + 1
    
    If UCase(lemparformcc) <> 1 Then
        FrmCC_Colection.Label12.Caption = i_monitoring_activity
        If bcp = False Then
            If FrmCC_Colection.Label12.Caption > 180 Then
                If UCase(MDIForm1.Text2.text) = "AGENT" Then
                    'M_OBJCONN.execute "UPDATE usertbl SET f_blok='1',alasan_blok='Tidak melakukan aktifitas 3 menit(2)' WHERE userid='" & Trim(MDIForm1.Text1.text) & "'"
                   ' MsgBox "Akun anda di blok, karena tidak melakukan aktivitas selama lebih dari 3 menit. oleh SPV/Admin! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
                   ' Call logwktcti("terblok timer2")
                   ' Call set_count_ol
                   ' Call offsesilogin(MDIForm1.Text1.text)
                   ' End
                End If
            End If
        End If
    ElseIf UCase(lemparformcc) = 1 Then
        frmCC_Colection2.Label12.Caption = i_monitoring_activity
        If bcp = False Then
            If frmCC_Colection2.Label12.Caption > 180 Then
                If UCase(MDIForm1.Text2.text) = "AGENT" Then
                    'M_OBJCONN.execute "UPDATE usertbl SET f_blok='1',alasan_blok='Tidak melakukan aktifitas 3 menit(2)' WHERE userid='" & Trim(MDIForm1.Text1.text) & "'"
                   ' MsgBox "Akun anda di blok, karena tidak melakukan aktivitas selama lebih dari 3 menit. oleh SPV/Admin! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
                   ' Call logwktcti("terblok timer2")
                   ' Call set_count_ol
                   ' Call offsesilogin(MDIForm1.Text1.text)
                   ' End
                End If
            End If
        End If
    End If
    
    'If TxtStatus.Text <> "FEEDBACKbusy" Then
'        If FrmCC_Colection.Label12.Caption > 180 Or frmCC_Colection2.Label12.Caption > 180 Then
'            If UCase(MDIForm1.Text2.Text) = "AGENT" Then
'                M_OBJCONN.Execute "UPDATE usertbl SET f_blok='1' WHERE userid='" & Trim(MDIForm1.Text1.Text) & "'"
'                MsgBox "Akun anda di blok, karena tidak melakukan aktivitas selama lebih dari 3 menit. oleh SPV/Admin! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
'                Call logwktcti("terblok timer2")
'                Call set_count_ol
'                Call offsesilogin
'                End
'            End If
'        End If
    'End If
End Sub

Sub KEDAPKEDIP()
    If Label4.Visible = True Then
        Label4.Visible = False
    ElseIf Label4.Visible = False Then
        Label4.Visible = True
    End If
End Sub

Private Sub Timer3_Timer()
    Call KEDAPKEDIP
    'Label3.Caption = Now()
    Dim M_DATA As New CLS_FRMCUST_CC_MGM
    Dim CMDSQL As String
    Dim n As String
    Dim tglnow As Date
    'tglnow=format(
    'Label4.Caption = Now()
    
    SqlWaktu = "select now()"
        Set m_waktuserver = New ADODB.Recordset
        m_waktuserver.CursorLocation = adUseClient
        m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    n = Format(m_waktuserver(0), "hh:mm:ss")
    
    If n = "22:00:00" Then
        Label3.Caption = "haiii"
    
        'Otomatis BP
        CMDSQL = "update mgm SET LASTSTATUS=KETHSLKERJA,KETHSLKERJA='BP-BROKEN PROMISE',F_CEK='BP-',REMARKS = 'BP-BROKEN PROMISE-Auto',RECSTATUS='C',OTO='Y',TGLSTATUS='" & Format(Now, "yyyy/mm/dd") & "'"
        CMDSQL = CMDSQL + "where custid in (select custid from vwptp1 "
        CMDSQL = CMDSQL + "where datediff(day,promisedate,getdate())>7 and custid not in ( "
        CMDSQL = CMDSQL + "select distinct custid from tbllunas)) And F_CEK like '%PTP%'"
        M_OBJCONN.execute CMDSQL
    
        'Otomatis POP
        CMDSQL = "update mgm SET"
        CMDSQL = CMDSQL + " LASTSTATUS=KETHSLKErJA,KETHSLKErJA='POP-PROGRESS OF PAYMENT',F_CEK='POP',rEMArKS = 'POP-PROGRESS OF PAYMENT-Auto',RECSTATUS='C',OTO='Y',TGLSTATUS='" & Format(Now, "yyyy/mm/dd") & "'"
        CMDSQL = CMDSQL + " where custid in ("
        CMDSQL = CMDSQL + " select distinct custid from tbllunas)"
        CMDSQL = CMDSQL + " And F_CEK<>'POP' AND F_CEK='PTP'"
        M_OBJCONN.execute CMDSQL
    
        '
        CMDSQL = "SELECT CUSTID,AGENT,REMARKS,NEXTACT,F_CEK,Statuscall,MOBILENO from mgm where OTO='Y'"
        Dim ds As ADODB.Recordset
        Set ds = New ADODB.Recordset
        ds.CursorLocation = adUseClient
        ds.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If ds.EOF And ds.BOF Then
        Else
            Do While Not ds.EOF
                M_DATA.ADD_HISTORY_OTO M_OBJCONN, ds!CustId, Format(Now, "yyyy/mm/dd hh:mm:dd"), Time, ds!agent, "COLLECTION", IIf(IsNull(ds!Remarks), "", ds!Remarks), "", IIf(IsNull(ds!NEXTACT), "", ds!NEXTACT), "", IIf(IsNull(ds!F_CEK), "", ds!F_CEK), IIf(IsNull(ds!statuscall), "", ds!statuscall), IIf(IsNull(ds!MOBILENO), "", ds!MOBILENO)
                ds.MoveNext
            Loop
        End If
    
        CMDSQL = "update mgm SET OTO=''"
        M_OBJCONN.execute CMDSQL
    End If
End Sub

Private Sub Timer4_Timer()
    Dim tglserver As String
    Dim TGLCLICK As String
    Dim ListItem As ListItem
    Dim ConnPTP As New ADODB.Connection
    Dim M_objrs As New ADODB.Recordset
    Dim cmdsql3 As String
    
    If shedulePTP = True Then
        'ngak ada kegiatan
    Else
    ConnPTP.Open CMDSQLOPEN
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    cmdsql3 = "select custid,name,tdbDatePTP from mgm where TdbDatePTP = '" + Format((Now + 7), "yyyy/mm/dd") + "' and agent ='" + MDIForm1.Text1.text + "'"
    'cmdsql3 = "select * from mgm where TGLINCOMING = '" + Format((MDIForm1.TDBDate1.Value + 7), "yyyy/mm/dd") + "'"
    M_objrs.Open cmdsql3, ConnPTP, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount <> 0 Then
        LstGrade.ListItems.clear
    End If
    While Not M_objrs.EOF
        Set ListItem = LstGrade.ListItems.ADD(, , M_objrs.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("NAME")), "", M_objrs("NAME"))
        ListItem.SubItems(3) = Format(IIf(IsNull(M_objrs("TdbDatePTP")), "", M_objrs("TdbDatePTP")), "yyyy/mm/dd hh:nn")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    If M_objrs.RecordCount <> 0 Then
        MsgBox "You Got Schedule PTP to Follow Up", vbInformation + vbOKOnly, "Aplikasi"
        shedulePTP = True
    End If

'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    cmdsql3 = "select custid, name, TdbDatePTP from mgm where TdbDatePTP = '" + Format((MDIForm1.TDBDate1.Value + 1), "yyyy/mm/dd") + "' and agent ='" + MDIForm1.Text1.Text + "'"
'    'cmdsql3 = "select * from mgm where TGLINCOMING = '" + Format((MDIForm1.TDBDate1.Value + 1), "yyyy/mm/dd") + "'"
'    m_objrs.Open cmdsql3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'    '    LstGrade.ListItems.Clear
'    End If
'    While Not m_objrs.EOF
'        Set listitem = LstGrade.ListItems.ADD(, , m_objrs.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
'        listitem.SubItems(3) = Format(IIf(IsNull(m_objrs("TdbDatePTP")), "", m_objrs("TdbDatePTP")), "yyyy/mm/dd hh:nn")
'        m_objrs.MoveNext
'    Wend
'    If m_objrs.RecordCount <> 0 Then
'        MsgBox "You Got Schedule PTP to Follow Up", vbInformation + vbOKOnly, "Aplikasi"
'        shedulePTP = True
'    End If
'End If
'Set m_objrs = Nothing

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    cmdsql3 = "select CUSTID, NAME, NEXTACTDATE from mgm where NEXTACTDATE BETWEEN '" + Format((Now), "yyyy/mm/dd") & " 00:00" + "' and '" + Format((Now), "yyyy/mm/dd") & " 23:59" + "' and agent ='" + MDIForm1.Text1.text + "'"
    'cmdsql3 = "select * from mgm where NEXTACTDATE BETWEEN '" + Format((MDIForm1.TDBDate1.Value), "yyyy/mm/dd") & " 00:00" + "' and '" + Format((MDIForm1.TDBDate1.Value), "yyyy/mm/dd") & " 23:59" + "'"
    M_objrs.Open cmdsql3, ConnPTP, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount <> 0 Then
        LstGrade.ListItems.clear
    End If
    While Not M_objrs.EOF
        Set ListItem = LstGrade.ListItems.ADD(, , M_objrs.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("NAME")), "", M_objrs("NAME"))
        ListItem.SubItems(3) = Format(IIf(IsNull(M_objrs("NEXTACTDATE")), "", M_objrs("NEXTACTDATE")), "yyyy/mm/dd hh:nn")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    ConnPTP.Close
    Set ConnPTP = Nothing
    End If
End Sub

Public Sub ActionCTI(Nilai As String)
    'ParameterCTI = Nilai & " /r/n"
    ParameterCTI = Nilai
    TimerCTI.Enabled = True
End Sub

Private Sub Timer5_Timer()
    Dim kalimat2 As String
         
    count_timer_detik = count_timer_detik + 1
         
    If PANJANG = COUNTER Then
        COUNTER = 0
        DoEvents
    Else
        '      Kalimat1 = Lbltargetspv.Caption
        COUNTER = COUNTER + 1
        DoEvents
        Lbltargetspv.Caption = TulisJalan(COUNTER, Kalimat1, 200)
    End If
    
    If count_timer_detik = 9 Then
        count_timer_detik = 0
        Label_OL_count = Label_OL_count + 1
    End If

End Sub

Private Sub Timer7_Timer()
    Dim dur_blok As Integer
    
    signtimer7 = True
    
    If UCase(MDIForm1.Text2.text) <> "AGENT" Then
        main_timer_activity = 0
        Timer7.Enabled = False
        Exit Sub
    End If
    
    waktu_finish = waktu_server_sekarang
    dur_blok = DateDiff("s", waktu_start, waktu_finish)
    
    DoEvents
    main_timer_activity = main_timer_activity + 1
    lbl_timer_activity = main_timer_activity
    'lbl_timer_activity.Caption = dur_blok
    lbl_timer_activity.Caption = main_timer_activity
    'If main_timer_activity > 180 Or dur_blok > 180 Then
    If lbl_timer_activity.Caption > 180 Then
        If UCase(MDIForm1.Text2.text) = "AGENT" Then
          '  M_OBJCONN.execute "UPDATE usertbl SET f_blok='1',alasan_blok='Tidak melakukan aktifitas 3 menit(7)' WHERE userid='" & Trim(MDIForm1.Text1.text) & "'"
            'MsgBox "Akun anda di blok, karena tidak melakukan aktivitas selama lebih dari 3 menit. oleh SPV/Admin! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
            'Call logwktcti("terblok timer7")
           ' Call set_count_ol
           ' Call offsesilogin(MDIForm1.Text1.text)
           ' End
        End If
    End If
End Sub

Private Sub Timer8_Timer()
'jejaktian===========================================================================
    Dim ConMSG As New ADODB.Connection
    Dim cmdsqlnew As String
    Dim cmdsql3 As String
    Dim M_objrs, m_objrs1 As New ADODB.Recordset
    Dim CMDSQL As String
    
    'On Error GoTo SALAH
    
    ConMSG.Open CMDSQLOPEN
    'M_Objrs.CursorLocation = adUseClient
    cmdsql3 = "select * from tampungtransferdata where y_n = 0 and tujapproval = '" + Text1.text + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql3, ConMSG, adOpenDynamic, adLockOptimistic, adCmdText

    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If

    While Not M_objrs.EOF
        cmdsql3 = "select * from tampungtransferdata where y_n = 0 and tujapproval = '" + Text1.text + "'"
        Set m_objrs1 = New ADODB.Recordset
        m_objrs1.CursorLocation = adUseClient
        m_objrs1.Open cmdsql3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
            If m_objrs1.RecordCount = 0 Then
                Set m_objrs1 = Nothing
                Exit Sub
            Else
                Formapptransferdata.lblpemohon.Caption = m_objrs1!pengupload
                Formapptransferdata.Label2.Caption = m_objrs1!pengupload
                Formapptransferdata.Label3.Caption = m_objrs1!tujapproval
                If m_objrs1.RecordCount <> 0 Then
                'On Error GoTo SALAH
                If Text1.text = m_objrs1!tujapproval Then
                    Formapptransferdata.Show vbModal
                Else
                    Exit Sub
                End If
            End If
            End If
    M_objrs.MoveNext
    Wend
            
    Exit Sub
'SALAH:
'bikin error
    'Formapptransferdata.Hide
    'MsgBox "Ada error : " & err.Description
End Sub

Private Sub Timer9_Timer()
    If bcp = False Then
        If MDIForm1.Text2.text <> "Agent" And MDIForm1.Text2.text <> "TeamLeader" Then
            Timer9.Enabled = True
            q = "SELECT * FROM tblvalidtospv"
            Set r = New ADODB.Recordset
            r.CursorLocation = adUseClient
            r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
            mntools.Caption = "&Tools "
            mnappvp.Caption = "Approve Valid Phone "
            If r.RecordCount <> 0 Then
                mntools.Caption = mntools.Caption & "(" & r.RecordCount & ")"
                mnappvp.Caption = mnappvp.Caption & "(" & r.RecordCount & ")"
            End If
        Else
            Timer9.Enabled = False
        End If
    End If
End Sub

Private Sub TimerAutoDialer_Timer()
On Error Resume Next
autodialer.Autdialer_CekON (MDIForm1.Text1.text)
If AutoDialerON = True Then
    VIEW_MGMDATA.Show
    If AutoDialerHangup = True Or AutoDialerBreak = False Then
       WaitSecs 5
       Autodialer_Calling (MDIForm1.Text1.text)
    Else
     
    End If
Else

End If

End Sub

Private Sub TimerBlink_Timer()
    If Label9.ForeColor = vbBlack Then
        Label9.ForeColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        Label9.ForeColor = vbBlack
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
        KelapKelip = 0
        'WaitSecs (3)
        'TimerBlink.Enabled = False
    End If
End Sub

Private Sub TimerCTI_Timer()
    Select Case WskCTI.state
        Case 9, 8, 1
            Debug.Print WskCTI.state
            WskCTI.Close
            WskCTI.RemoteHost = "127.0.0.1"
            'buat connect ke chromium
            'WskCTI.RemotePort = 2121
            'buat connect ke cti
            WskCTI.RemotePort = 18000
            WskCTI.Connect
        Case 6
            Debug.Print WskCTI.state
        Case 7
            Debug.Print WskCTI.state
            If Len(ParameterCTI) > 2 Then
                WskCTI.SendData ParameterCTI + vbCrLf
                'MsgBox ParameterCTI
            End If
            Debug.Print ParameterCTI
            TimerCTI.Enabled = False
        Case 0
            Debug.Print WskCTI.state
            WskCTI.RemoteHost = "127.0.0.1"
            'buat connect ke chromium
            'WskCTI.RemotePort = 2121
            'buat connect ke cti
            WskCTI.RemotePort = 18000
            WskCTI.Connect
        Case Else
    End Select
End Sub

Private Sub TimerRequest_Timer()
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Or _
       UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.Text2.text) = "ADMIN" Or _
       UCase(MDIForm1.Text2.text) = "ADMINISTRATOR" Then
       Call CekReqNumber
    Else
        TimerRequest.Enabled = False
    End If
End Sub

Private Sub CekReqNumber()
'@@11092012 - DinonAktifkan
'    Dim CMDSQL As String
'    Dim M_OBJRS As ADODB.Recordset
'
'    'cek status f_req_number
'    CMDSQL = "select * from usertbl where userid='"
'    CMDSQL = CMDSQL + MDIForm1.Text1.Text + "'"
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_OBJRS.RecordCount > 0 Then
'        If M_OBJRS("f_req_number") = "1" Then
'            On Error GoTo salah
'            FrmPemberitahuan.Show vbModal
'        End If
'    End If
'
'    Set M_OBJRS = Nothing
'    Exit Sub
'salah:
'    FrmPemberitahuan.Hide
'    MsgBox "Ada error :" & Err.Description
End Sub


Private Sub TimerTandaReq_Timer()
    If ShapeReq.FillColor = vbBlack Then
        ShapeReq.FillColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        ShapeReq.FillColor = vbBlack
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
        KelapKelip = 0
        WaitSecs (3)
        ShapeReq.FillColor = vbBlack
        TimerTandaReq.Enabled = False
    End If
End Sub

Private Sub transfer_data_Click()
    'formtransferdata.Show vbModal
    frm_transferdata2.Show vbModal
End Sub

Private Sub update_status_ptp_Click()
    Formupdate_status_ptp.Show vbModal
End Sub

Private Sub upload_fresh_wo_Click()
    Form_upload_fresh_wo.Show vbModal
End Sub

Private Sub VSMS_Click()
    If UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Load Frm_verify
    Frm_verify.Show vbModal
End Sub

Private Sub WskCTI_DataArrival(ByVal bytesTotal As Long)
    Dim StrMsgDrCti As String
    Dim lStrMsgDrCti As String
    Dim strnew As String
    
    
    'FrmCC_Colection.Refresh
    
    'get data balik dari cti
    'StrMsgDrCti = "FEEDBACKhangup"
    WskCTI.GetData StrMsgDrCti, vbString
    Debug.Print StrMsgDrCti
    On Error GoTo bawah:
    
    'uniqpublic = ""
    
    'sementara ditutup menunggu dccs baru 23januari 2018
    'open 7febuari2018
    Dim output() As String
    output = Split(StrMsgDrCti, "|")
    strnew = output(0)
    If strnew = "FEEDBACKhangup" Or strnew = "FEEDBACKinitiated" Then
    
    'If Len(StrMsgDrCti) > 50 Then
        uniqpublic = output(2)
        txtuniqueid.text = output(2)
        'MsgBox output(2)
    Else
    End If
    
    If strnew = "FEEDBACKinitiated" Then
         WsckCti_initiated = "FEEDBACKinitiated"
    End If
bawah:
    
    '---------------------------------------------
    
    
    If Len(StrMsgDrCti) > 1 Then
        'FrmSoftPhone.Caption = StrMsgDrCti
        lStrMsgDrCti = Left(StrMsgDrCti, Len(StrMsgDrCti) - 2)
        TxtStatus.text = StrMsgDrCti

        If InStr(lStrMsgDrCti, "connected..") Or strnew = "FEEDBACKconnected" Then
            WsckCti_connected = "FEEDBACKconnected"
            cti_get = "FEEDBACKconnected"
            Obelisk = True
        End If
        
        If InStr(lStrMsgDrCti, "FEEDBACKprogressing") Then
            WsckCti_connected = "FEEDBACKconnected"
            
        End If
        
        'TxtStatus.Text = Mid(lStrMsgDrCti, InStr(1, lStrMsgDrCti, "K", vbTextCompare) + 1, Len(lStrMsgDrCti) - InStr(1, lStrMsgDrCti, "K", vbTextCompare))
    
        'Debug.Print Mid(lStrMsgDrCti, InStr(1, lStrMsgDrCti, "K", vbTextCompare) + 1, Len(lStrMsgDrCti) - InStr(1, lStrMsgDrCti, "K", vbTextCompare))
        'Debug.Print Len(Mid(lStrMsgDrCti, InStr(1, lStrMsgDrCti, "K", vbTextCompare) + 1, Len(lStrMsgDrCti) - InStr(1, lStrMsgDrCti, "K", vbTextCompare)))
      '  Debug.Print "text - " & TxtStatus.Text
      
'      If TxtStatus.Text = "ringing" Then
'        FrmAnswerCall.Show vbModal
'      Else
'        If InStr(lStrMsgDrCti, "ringing") Then
'          FrmAnswerCall.Show vbModal
'        End If
'      End If
'
      If InStr(lStrMsgDrCti, "free") Then
        'hang up
'            Call savecall
'            FBILL.Timer6.Enabled = False
'            Unload FBILL
 '           Unload FrmAnswerCall
      End If
'      If InStr(lStrMsgDrCti, "busy") And rounding <> 0 Then
'        If CmbNo.Text = "108" Or CmbNo.Text = "147" Or CmbNo.Text = "109" Then
'        Else
'            'Di angkat
'             FBILL.Timer6.Enabled = True
'             FBILL.Show
'        End If
'      End If

        ' 14 Agustus 2014 ----- 5x Call
        'If InStr(lStrMsgDrCti, "FEEDBACKbusy") Then
            
        'End If
        
        'lStrMsgDrCti = "FEEDBACKprogressing"
        'If InStr(lStrMsgDrCti, "FEEDBACKbusy") then YANG LAMA

        If (lStrMsgDrCti Like "*FEEDBACKbusy*") Or (lStrMsgDrCti Like "*FEEDBACKbu*") Or (lStrMsgDrCti Like "*FEEDBACKprogressing*") Or (lStrMsgDrCti Like "*FEEDBACKprogress*") Or strnew = "FEEDBACKbusy" Or strnew = "FEEDBACKprogressing" Then
            'STATUS = "FEEDBACKbusy"
            WsckCti_busy = "FEEDBACKbusy"
            i_monitoring_activity = 0
            Timer2.Enabled = False
'            i_monitoring_activity_2 = 0
'            Timer9.Enabled = False
            Call logwktcti("reset waktu")
            'waktu_mulai_ngitung = waktu_server_sekarang
            'waktu_mulai_ngitung = waktu_server_sekarang
            'Call insertlogcti(STATUS)
            
        End If
        
        If InStr(lStrMsgDrCti, "FEEDBACKhangup") Or InStr(lStrMsgDrCti, "FEEDBACKprogressingFEEDBACKhangup") Or (lStrMsgDrCti Like "*FEEDBACKhangup*") Or (lStrMsgDrCti Like "*FEEDBACKhang*") Or strnew = "FEEDBACKhangup" Then
            'STATUS = "FEEDBACKHANGUP"
            WsckCti_hangup = "FEEDBACKhangup"

            Timer2.Enabled = True
            i_monitoring_activity = 0
'            Timer9.Enabled = True
'            i_monitoring_activity_2 = 0
            waktu_mulai_ngitung = waktu_server_sekarang
            waktu_selesai_ngitung = waktu_server_sekarang
            'Call insertlogcti(STATUS)
            FrmCC_Colection.SSCommand1_Click (1)
        End If
        
        If InStr(lStrMsgDrCti, "FEEDBACKfree") Or InStr(lStrMsgDrCti, "FEEDBACKready") Or strnew = "FEEDBACKready" Then
            'STATUS = "FEEDBACKFREE"
            Timer2.Enabled = True
            i_monitoring_activity = 0
'            Timer9.Enabled = True
'            i_monitoring_activity_2 = 0
            waktu_mulai_ngitung = waktu_server_sekarang
            waktu_selesai_ngitung = waktu_server_sekarang
            'Call insertlogcti(STATUS)
        End If
        
        If InStr(lStrMsgDrCti, "FEEDBACKcdr") Then
            Dim arr_str() As String
            Dim str() As String
            Dim struniqueid As String
            txtdurasi.text = "0"
            arr_str = Split(lStrMsgDrCti, "|")
            If arr_str(0) = "FEEDBACKcdr" Then
            actiondccs = arr_str(7) ' unique_id
            struniqueid = arr_str(7)
            str = Split(struniqueid, vbCrLf)
            'txtuniqueid.text = str(0)
            txtdurasi.text = arr_str(6)
            End If
            
            
        End If
        
        
        
        ' Untuk Form sms 02/07/2013 By Izuddin
        If open_sms Then
            Timer2.Enabled = False
            'Timer9.Enabled = False
        End If
        Call logwktcti(waktu_mulai_ngitung & " " & lStrMsgDrCti)
    End If
    
End Sub

Private Sub WskCTI_SendComplete()
    Debug.Print "OK send"
End Sub

Private Sub WskCTI_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Debug.Print "Sending" & CStr(bytesSent) & "/" & CStr(bytesRemaining)
End Sub

'softphone control
Private Sub CmdACW_Click()
    MDIForm1.ActionCTI ("ACW")
End Sub

Private Sub CmdAUX_Click()
    MDIForm1.ActionCTI ("AUX")
End Sub

Private Sub CmdCall_Click()
'    If CmbNo.Text = "108" Then
'        MDIForm1.ActionCTI ("DIAL|" & MDIForm1.TxtAuthPrefix.Text & GetNumber(108))
'    ElseIf CmbNo.Text = "109" Then
'        MDIForm1.ActionCTI ("DIAL|" & MDIForm1.TxtAuthPrefix.Text & GetNumber(109))
'    ElseIf CmbNo.Text = "147" Then
'        MDIForm1.ActionCTI ("DIAL|" & MDIForm1.TxtAuthPrefix.Text & GetNumber(147))
'    Else
        'MDIForm1.ActionCTI ("DIAL|" + CmbNo)
'    End If

    If CmbNo.text = "108" Then
        MDIForm1.ActionCTI ("DIAL|268" & GetNumber(108))
    Else
        If Len(CmbNo.text) >= 6 And Len(CmbNo.text) <= 7 Then
            If Left(CmbNo.text, 2) <> "08" And Right(CmbNo.text, 3) = "108" Then
                MDIForm1.ActionCTI ("DIAL|268" & GetNumber(CmbNo.text))
            End If
        End If
    End If
    
    If CmbNo.text = "109" Then
        MDIForm1.ActionCTI ("DIAL|268" & GetNumber(109))
    Else
        If Len(CmbNo.text) >= 6 And Len(CmbNo.text) <= 7 Then
            If Left(CmbNo.text, 2) <> "08" And Right(CmbNo.text, 3) = "109" Then
                MDIForm1.ActionCTI ("DIAL|268" & GetNumber(CmbNo.text))
            End If
        End If
    End If
    
    If CmbNo.text = "147" Then
        MDIForm1.ActionCTI ("DIAL|268" & GetNumber(147))
    Else
        If Len(CmbNo.text) >= 6 And Len(CmbNo.text) <= 7 Then
            If Left(CmbNo.text, 2) <> "08" And Right(CmbNo.text, 3) = "147" Then
                MDIForm1.ActionCTI ("DIAL|268" & GetNumber(CmbNo.text))
            End If
        End If
    End If
    
    If CmbNo.text = "109" Or CmbNo.text = "108" Or CmbNo.text = "147" Then
    Else
        MDIForm1.ActionCTI ("DIAL|" + CmbNo)
    End If
End Sub

Private Sub CmdConference_Click()
    MsgBox "Conference"
End Sub

Private Sub CmdHangUp_Click()
    MDIForm1.ActionCTI ("HANGUP")
End Sub

Private Sub CmdLogin_Click()
    MsgBox "Login"
End Sub

Private Sub Cmdready_Click()
    MDIForm1.ActionCTI ("READY")
End Sub

Private Sub CmdOutbound_Click()
    MDIForm1.ActionCTI ("NOTREADY")
End Sub

Private Sub CmdTransfer_Click()
    If CmbNo.text = "" Then
    Else
        MDIForm1.ActionCTI ("TRANSFER" + CmbNo)
    End If
End Sub

Private Sub CmdBintang_Click()
    CmbNo.text = CmbNo.text + "*"
End Sub

Private Sub CmdCancel_Click()
    CmbNo.text = ""
End Sub

Private Sub CmdNo_Click(Index As Integer)
    CmbNo.text = CmbNo.text + CStr(Index)
End Sub

Private Sub CmdPager_Click()
    CmbNo.text = CmbNo.text + "#"
End Sub


Private Sub Cmddtmf_Click()
    If CmbNo.text = "" Then
    Else
        MDIForm1.ActionCTI ("DTMF" + CmbNo)
    End If
End Sub

Private Sub PromiseToPay()
    LstGrade.ColumnHeaders.ADD 1, , "No", 3 * 120
    LstGrade.ColumnHeaders.ADD 2, , "Cust ID", 10 * 120
    LstGrade.ColumnHeaders.ADD 3, , "Status", 10 * 120
    LstGrade.ColumnHeaders.ADD 4, , "Agent", 10 * 120
End Sub

Private Sub HeaderInformation()
    LstInformation.ColumnHeaders.ADD 1, , "Description", 20 * 120
    LstInformation.ColumnHeaders.ADD 2, , "No", 1
    LstInformation.ColumnHeaders.ADD 3, , "Lokasi", 1
End Sub

Private Sub CmdAccept_Click()
    MDIForm1.ActionCTI ("DTMF")
End Sub

Private Sub CmdOutgoing_Click()
    MDIForm1.ActionCTI ("OUTGOING")
End Sub

Private Sub LstInformation_DblClick()
    If LstInformation.ListItems.Count = 0 Then
        Exit Sub
    End If
    If StartMeUp(LstInformation.SelectedItem.SubItems(2)) <= 32 Then
       MsgBox "File Tidak Ditemukan", vbOKOnly + vbCritical, "Pemberitahuan"
       
    Else
    SSTab1.Tab = 0
    End If
End Sub

Private Sub LstDataInformation()
    Dim ListItem As ListItem
    Dim ssql As String
    Set M_LOGINRS = New ADODB.Recordset
    M_LOGINRS.CursorLocation = adUseClient

    ssql = "SELECT ExpiryDate, Description, id, Direktori FROM tblinformationlokasi " & _
           "ORDER BY Description"
    M_LOGINRS.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    While Not M_LOGINRS.EOF
    If Format(M_LOGINRS!ExpiryDate, "yyyy/mm/dd") > Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") Then
        Set ListItem = MDIForm1.LstInformation.ListItems.ADD(, , IIf(IsNull(M_LOGINRS("Description")), "", M_LOGINRS("Description")))
            ListItem.SubItems(1) = IIf(IsNull(M_LOGINRS("id")), "", M_LOGINRS("id"))
            ListItem.SubItems(2) = IIf(IsNull(M_LOGINRS("Direktori")), "", M_LOGINRS("Direktori"))
    End If
        M_LOGINRS.MoveNext
    Wend

    Set M_LOGINRS = Nothing
End Sub

Function ReplaceFirstInstance(SourceString, _
    Searchstring, Replacestring)
    Dim StartLoc
    Dim FoundLoc
    If StartLoc = 0 Then StartLoc = 1
    FoundLoc = InStr(StartLoc, SourceString, Searchstring) '*
    If FoundLoc <> 0 And FoundLoc < 2 Then
       ReplaceFirstInstance = Left(SourceString, FoundLoc - 1) & Replacestring & Right(SourceString, Len(SourceString) - (FoundLoc - 1) - Len(Searchstring))
       StartLoc = FoundLoc + Len(Replacestring)
    ElseIf FoundLoc > 1 Then
        ReplaceFirstInstance = Replacestring & "21" & SourceString
    Else
        StartLoc = 1
        ReplaceFirstInstance = SourceString
    End If
End Function

Function FindReplace(SourceString, Searchstring, Replacestring) As String
    Dim tmpString1
    Dim tmpString2
    tmpString1 = SourceString
 
    tmpString2 = tmpString1
    tmpString1 = ReplaceFirstInstance(tmpString1, _
                 Searchstring, Replacestring)
    
    FindReplace = tmpString1
End Function

'@@ 15-12-2010 buat timer lock data
Private Sub Timer_stopwatch_Timer()
    
    'Tambah dengan satu untuk total sepersepuluh detik.
    'Kita mengeset interval Timer menjadi 10, jadi
    'setiap sepersepuluh detik prosedur ini akan
    'dieksekusi
    TotalTenthDetik = TotalTenthDetik + 1
    'Jika TotalTenthSeconds = 10,
    'set kembali menjadi 0.
    TenthDetik = TotalTenthDetik Mod 10
    '10 kali sepersepuluh detik sama dengan 1 detik.
    'int - akan mengembalikan bilangan integer (bulat)
    'dari pecahan 'Contoh: Int(0.9) = 0 menghasilkan 0
    TotalDetik = Int(TotalTenthDetik / 10)
    'Jika variabel Seconds = 60, set kembali menjadi 0
    Detik = TotalDetik Mod 60
    If Len(Detik) = 1 Then
       Detik = "0" & Detik  'Agar selalu dalam dua
                            'digit
    End If
    Menit = Int(TotalDetik / 60) Mod 60
    If Len(Menit) = 1 Then
       Menit = "0" & Menit    'Agar selalu dalam dua
                          'digit
    End If
    JAM = Int(TotalDetik / 3600)
    If JAM < 9 Then
       jam1 = "0" & JAM       'Agar selalu dalam dua'digit
    End If
    'Tampilkan hasilnya di Lblwaktu (update terus Lblwaktu)
    LblWaktu.Caption = jam1 & ":" & Menit & ":" & Detik & ":" & TenthDetik & ""
             
    If LblWaktu.Caption = TxtWaktuRefresh.text & ":0" Then
        DoEvents
        
        If Format(Now, "hh:mm:ss") > CDate(#9:00:00 PM#) Then
            GoTo selanjutnya
        End If
        
        If UCase(Text2.text) = "TEAMLEADER" Then
             'MsgBox "Ok"
             Call LockDataAuto
        End If
       
selanjutnya:
        'Memulai atau menghentikan timer kembali
         'Timer_stopwatch.Enabled = Not Timer1.Enabled
         'Inisialisasi total sepersepuluh detik
         TotalTenthDetik = -1
         'Aktifkan timer
         Timer_stopwatch.Enabled = True
       
    End If


End Sub

Private Sub LockDataAuto()
        '@@ Awal 061110 cek lock account sesuai settingan timer
        Dim m_objrsTemp As ADODB.Recordset
        Dim M_ObjrsWaktuServer As ADODB.Recordset
        Dim m_objrsCurrent As ADODB.Recordset
        
        
        Dim cmdsqlserver As String
        Dim WaktuServer As Date
        Dim WaktuAkhirCurrent As Date
        
        'ambil waktu server
        cmdsqlserver = "select now() as WaktuServer "
        Set M_ObjrsWaktuServer = New ADODB.Recordset
        M_ObjrsWaktuServer.CursorLocation = adUseClient
    
        M_ObjrsWaktuServer.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        WaktuServer = Format(M_ObjrsWaktuServer(0), "mm-dd-yyyy hh:mm")
        Set M_ObjrsWaktuServer = Nothing
        
        'Cek lock account yang sedang berjalan
        cmdsqlserver = "select * from tbltemplockacc_current "
        Set m_objrsCurrent = New ADODB.Recordset
        m_objrsCurrent.CursorLocation = adUseClient
        m_objrsCurrent.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If m_objrsCurrent.RecordCount <> 0 Then
            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
        Else
            GoTo lockdata
        End If
        
        While Not m_objrsCurrent.EOF
            
            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
        
            If WaktuAkhirCurrent <= WaktuServer Then
                'Cek dulu apakah ada user yang sedang mereset data
                If Trim(m_objrsCurrent("f_locked")) = "2" Then
                    GoTo KeluarLockAutoTL
                End If
                
                'update dulu status lock yang sedang berakhir, supaya agent lain ga ikut ngereset
                cmdsqlserver = "update tbltemplockacc_current set f_locked='2' where id='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
                M_OBJCONN.execute cmdsqlserver
            
                'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
                cmdsqlserver = "update usertbl set dilockoleh='ClearByAutomatic',"
                cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
                cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null,f_pesanlockauto=null,f_idsessstart=null,f_pesanresetauto='1',f_idsessend='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "' "
                'Buat ambil kondisi agent yang sedang di lock
                If Trim(m_objrsCurrent("account_lock")) = "ALL" Then
                    cmdsqlserver = cmdsqlserver + " where usertype='1' "
                ElseIf Left(Trim(m_objrsCurrent("account_lock")), 3) = "SPV" Then
                    cmdsqlserver = cmdsqlserver + " where spvcode='"
                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "' "
                Else
                    cmdsqlserver = cmdsqlserver + " where userid='"
                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "' "
                End If
'                cmdsqlserver = cmdsqlserver + " and f_idsessstart='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "' "
                M_OBJCONN.execute cmdsqlserver
                
                
                'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
                cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current "
                cmdsqlserver = cmdsqlserver + " where id='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
                M_OBJCONN.execute cmdsqlserver
                
                'Hapus data di tabel locktemp current
                cmdsqlserver = "delete from tbltemplockacc_current where id='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
                M_OBJCONN.execute cmdsqlserver
                
             End If
KeluarLockAutoTL:
                m_objrsCurrent.MoveNext
            Wend
            Set m_objrsCurrent = Nothing

            
       
        
        '=======
lockdata:
        'Setelah cek waktu lock yang habis, sekarang cek lock yg masih dalam antrian
        cmdsqlserver = "select * from tbltemplockacc where f_locked isnull order by start_lock asc "
        Set m_objrsTemp = New ADODB.Recordset
        m_objrsTemp.CursorLocation = adUseClient
        m_objrsTemp.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            'Cek ada ga data lock dalam antrian
            If m_objrsTemp.RecordCount <> 0 Then
                Dim WaktuAwal As Date
                Dim WaktuAkhir As Date
                
                While Not m_objrsTemp.EOF
                
                    WaktuAwal = Format(m_objrsTemp("start_lock"), "mm-dd-yyyy hh:mm")
                    WaktuAkhir = Format(m_objrsTemp("end_lock"), "mm-dd-yyyy hh:mm")
                    
                    If (WaktuAwal <= WaktuServer) And (WaktuAkhir > WaktuServer) Then
                        'Cek apakah datanya sedang di lock sama agent lain?
                        If Trim(m_objrsTemp("f_locked")) = "1" Then
                            GoTo KeluarLockAutoTLLock
                        End If
                        
                        'update status  f_lockednya jadi 1, supaya ga di log sama agent lain
                        cmdsqlserver = "update tbltemplockacc set f_locked='1' where id='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
                        M_OBJCONN.execute cmdsqlserver
                        
                        'LAKUKAN LOCK DATA
                        Dim i As Integer
                       
                        a = Split(m_objrsTemp("script_lock"), "|")
                        
                        For i = LBound(a) + 1 To UBound(a) - 1
                            cmdsqlserver = Replace(a(i), "$", "'")
                            M_OBJCONN.execute cmdsqlserver
                        Next i
                        
                        'Pindahin dulu data di tabel current ke tabel log, terus data di tabel current dihapus
'                        cmdsqlserver = "insert into tbltemplockacc_current "
'                        cmdsqlserver = cmdsqlserver + " select * from tbltemplockacc_log"
'                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
                        
'                        cmdsqlserver = "delete from tbltemplockacc_current"
'                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
                        
                        'Pindahin data dari tabel temp lock ke tabel current log
                        cmdsqlserver = "insert into tbltemplockacc_current "
                        cmdsqlserver = cmdsqlserver + "select * from tbltemplockacc where id='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
                        M_OBJCONN.execute cmdsqlserver
                        
                        
                        
                       'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
                        cmdsqlserver = "update usertbl set f_pesanlockauto='1',f_idsessstart='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "' "
                        'Buat mengupdate pesan kondisi agent yang di lock
                        If Trim(m_objrsTemp("account_lock")) = "ALL" Then
                            cmdsqlserver = cmdsqlserver + " where usertype='1' "
                        ElseIf Left(Trim(m_objrsTemp("account_lock")), 3) = "SPV" Then
                            cmdsqlserver = cmdsqlserver + " where spvcode='"
                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
                        Else
                            cmdsqlserver = cmdsqlserver + " where userid='"
                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
                        End If
                        M_OBJCONN.execute cmdsqlserver
                        
                        'Hapus data di templock
                        cmdsqlserver = "delete from tbltemplockacc where id='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
                        M_OBJCONN.execute cmdsqlserver
                        
                        
                    End If
                   
KeluarLockAutoTLLock:
                    m_objrsTemp.MoveNext
               Wend

            End If
        
        Set m_objrsTemp = Nothing
      
      '@@ Akhir 061110 cek lock account sesuai settingan timer
End Sub

'@@ 14022011 ini buat cek sms
Private Sub CekSms()
    On Error Resume Next
    'Dim ConnPTP As New ADODB.Connection
    Dim M_objrs As New ADODB.Recordset
    Dim cmdsql34 As String
    Dim TELPo As String
    Dim codea As String
    
    If Left(Text1, 1) = "D" Or Text1 = "JOKO" Or Text1 = "SPV1" Or Left(Text1, 1) = "T" Then
        Select Case Text1.text
            Case "TL1"
                codea = "ACC1"
            Case "TL2"
                codea = "ACC2"
            Case "TL3"
                codea = "ACC3"
            Case "TL4"
                codea = "ACC4"
            Case "TL5"
                codea = "ACC5"
            Case "TL6"
                codea = "ACC6"
            Case "TL7"
                codea = "ACC7"
            Case "TL8"
                codea = "ACC8"
            Case "TL9"
                codea = "ACC9"
            Case "TL10"
                codea = "ACC10"
            Case Else
                codea = Text1.text
        End Select
    
        TELPo = "Select count(*) as banyak from inbox where sendernumber in ('a',"
    
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + codea + "'"
        M_objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
        If M_objrs.RecordCount = 0 Then
            Timer6.interval = 60000
            Exit Sub
        End If
    
        While Not M_objrs.EOF
            If Len(M_objrs("mobileno")) <> 0 Then
                satu = FindReplace(M_objrs("mobileno"), "0", "+62")
                TELPo = TELPo + "'" + satu + "',"
            Else
                TELPo = TELPo
            End If
    
            If Len(M_objrs("mobileno2")) <> 0 Then
                dua = FindReplace(M_objrs("mobileno2"), "0", "+62")
                TELPo = TELPo + "'" + dua + "',"
            Else
                TELPo = TELPo
            End If
    
            If Len(M_objrs("mobilenoadd1")) <> 0 Then
                tiga = FindReplace(M_objrs("mobilenoadd1"), "0", "+62")
                TELPo = TELPo + "'" + tiga + "',"
            Else
                TELPo = TELPo
            End If
            
            If Len(M_objrs("mobilenoadd2")) <> 0 Then
                empat = FindReplace(M_objrs("mobilenoadd2"), "0", "+62")
                TELPo = TELPo + "'" + empat + "',"
            Else
                TELPo = TELPo
            End If
        
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
    
        TELPo = Left(TELPo, Len(TELPo) - 1)
        Dim TELPo1
        Dim TELPo2
    
        TELPo1 = TELPo + ") and processed='f'"
        TELPo2 = TELPo + ") and processed='t'"
        
        If M_OBJCONN1.state = 1 Then
        M_objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            'On Error Resume Next
            LblJmlSmsBaru.Caption = M_objrs("banyak")
            Label9 = "SMS BARU " & LblJmlSmsBaru.Caption & " SMS"
            M_objrs.MoveNext
        Wend
    
     End If
     
        'JIKA ADA SMS BARU MASUK
        If Trim(Label9.Caption) = "SMS BARU 0 SMS" Then
            'MsgBox "Tidak ada sms baru!"
            TimerBlink.Enabled = False
            Label9.ForeColor = vbBlack
        Else
            If Trim(Label9.Caption) <> "" Then
                TimerBlink.Enabled = True
                MsgBox "Ada SMS BARU MASUK! Silahkan cek!", vbOKOnly + vbInformation, "Informasi"
            End If
        End If
    
        Set M_objrs = Nothing
    
    '-----------------------------
       If M_OBJCONN1.state = 1 Then
        M_objrs.Open TELPo2, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            Label10 = "SMS LAMA " & M_objrs("banyak") & " SMS"
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
        End If
        
    End If
    
    'MsgBox TELPo
     ' M_OBJCONN.Close
      'Set M_OBJCONN = Nothing
    Timer6.interval = 60000
End Sub

'@@06-04-2011, Tambahan jika sudah jam 11 Siang maka TL diingatkan untuk segera menarik report Contacto
Private Sub TimerTanda_Timer()
    If ShapeTanda.FillColor = vbBlack Then
        ShapeTanda.FillColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        ShapeTanda.FillColor = vbBlack
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
        KelapKelip = 0
        WaitSecs (3)
        'TimerBlink.Enabled = False
    End If
End Sub

Private Sub TimerWaktu_Timer()
    '@@06-04-2011 Jika yang login Agent, matikan timer waktu
    If UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        ShapeTanda.Visible = False
        TimerWaktu.Enabled = False
    End If
    LblWaktu.Caption = Format(Now, "hh:mm:ss")
'    If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Or _
'       UCase(Trim(MDIForm1.Text2.Text)) = "ADMIN" Or _
'       UCase(Trim(MDIForm1.Text2.Text)) = "ADMINISTRATOR" Then
'            If LblWaktu.Caption = "11:00:00" Then
'                TimerTanda.Enabled = True
'                MsgBox "Sudahkah anda menarik report Productivity? Tekan Ok untuk melihatnya!", vbOKOnly + vbInformation, "Informasi"
'                Call IsiAgentContactto
'                Call IsiContactto
'                Call IsiContacttoJmlAcc
'                WaitSecs (2)
'                FrmMgmReport.RPT.Reset
'                FrmMgmReport.RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'                FrmMgmReport.RPT.Formulas(2) = "@TglShow = totext('" + CStr(Format(Now, "dd-mm-yyyy") & " " & Format(Now, "hh:mm:ss")) + "')"
'                FrmMgmReport.RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Format(Now, "dd-mm-yyyy") & " " & Format(Now, "hh:mm:ss")) + "')"
'                FrmMgmReport.RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptContactto.rpt"
'                MDIForm1.TimerTanda.Enabled = False
'                MDIForm1.ShapeTanda.FillColor = vbBlack
'                Call SHOW_PRN
'            End If
'    End If
End Sub


'@@ 17-03-2011 Report Contactto
Private Sub IsiAgentContactto()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    
        CMDSQL = "select distinct u.spvcode as spv ,m.agent as agent"
        CMDSQL = CMDSQL + " from mgm as m, usertbl as u where "
        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
        CMDSQL = CMDSQL + " where spvcode between '"
        CMDSQL = CMDSQL + "SPV1" + "' and '"
        CMDSQL = CMDSQL + "SPV9" + "' and usertype='1') and date(m.tglcall) between '"
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' "
        CMDSQL = CMDSQL + " group by m.agent,u.spvcode "
        CMDSQL = CMDSQL + "order by u.spvcode,m.agent asc"
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    M_RPTCONN.execute "delete from TblRptContactto"
    If M_objrs.RecordCount > 0 Then
        'ProgressBar1.Max = M_OBJRS.RecordCount
        While Not M_objrs.EOF
            'ProgressBar1.Value = M_OBJRS.Bookmark
            CMDSQL = "insert into TblRptContactto (spvcode,agent) values ('"
            CMDSQL = CMDSQL + Trim(M_objrs("spv")) + "','"
            CMDSQL = CMDSQL + Trim(M_objrs("agent")) + "')"
             M_RPTCONN.execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub

Private Sub IsiContactto()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    
'        CMDSQL = "select u.spvcode as spv ,m.agent as agent,m.stscallwith as status,count(m.stscallwith) as jumlah "
'        CMDSQL = CMDSQL + " from mgm as m, usertbl as u where "
'        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
'        CMDSQL = CMDSQL + " where spvcode between '"
'        CMDSQL = CMDSQL + "SPV1" + "' and '"
'        CMDSQL = CMDSQL + "SPV9" + "' and usertype='1') and date(m.tglcall) between '"
'        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
'        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' "
'        CMDSQL = CMDSQL + " group by m.agent,m.stscallwith,u.spvcode "
'        CMDSQL = CMDSQL + "order by m.agent,m.stscallwith,u.spvcode asc"

        '@@01 Juni 2011 diubah querynya
        CMDSQL = "select u.spvcode as spv ,m.agent as agent,m.ststelpwith as status,count(m.ststelpwith) as jumlah "
        CMDSQL = CMDSQL + " from mgm_hst as m, usertbl as u where "
        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
        CMDSQL = CMDSQL + " where spvcode between '"
        CMDSQL = CMDSQL + "SPV1" + "' and '"
        CMDSQL = CMDSQL + "SPV9" + "' and usertype='1') and date(m.tgl) between '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' and "
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and "
        CMDSQL = CMDSQL + " m.ststelpwith in ('OTHER','CH','SPOUSE','PARENT')"
        CMDSQL = CMDSQL + " group by m.agent,m.ststelpwith,u.spvcode "
        CMDSQL = CMDSQL + "order by m.agent,m.ststelpwith,u.spvcode asc"

    
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        'ProgressBar1.Max = M_OBJRS.RecordCount
        While Not M_objrs.EOF
'             If Trim(m_objrs("status")) <> "SPOUSE" Or _
'                Trim(m_objrs("status")) <> "CONTACTED-CH" Or _
'                Trim(m_objrs("status")) <> "OTHER" Or _
'                Trim(m_objrs("status")) <> "PARENT" Or _
'                Trim(m_objrs("status")) <> "CH" Or _
'                Trim(m_objrs("status")) = "" Then
'                m_objrs.MoveNext
'             End If
            On Error Resume Next
            'ProgressBar1.Value = M_OBJRS.Bookmark
             CMDSQL = "update tblrptcontactto set ["
             CMDSQL = CMDSQL + Trim(Replace(M_objrs("status"), "/", "")) + "]='"
             CMDSQL = CMDSQL + CStr(M_objrs("jumlah")) + "' where spvcode='"
             CMDSQL = CMDSQL + Trim(M_objrs("spv")) + "' and agent='"
             CMDSQL = CMDSQL + Trim(M_objrs("agent")) + "'"
             M_RPTCONN.execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub
'@@01 Juni 2011
Private Sub IsiContacttoJmlAcc()
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs_rpt As ADODB.Recordset
    Dim CMDSQL As String
    
    
    'Ambil Data agentnya
    CMDSQL = "select spvcode,agent from tblrptcontactto order by spvcode,agent"
    Set m_objrs_rpt = New ADODB.Recordset
    m_objrs_rpt.CursorLocation = adUseClient
    m_objrs_rpt.Open CMDSQL, M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs_rpt.RecordCount > 0 Then
        
        While Not m_objrs_rpt.EOF
            
            CMDSQL = "select distinct custid from mgm_hst where agent='"
            CMDSQL = CMDSQL + Trim(m_objrs_rpt("agent")) + "' and date(tgl) between '"
            CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
            CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' "
            
            
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            'update data ke access
            CMDSQL = "update tblrptcontactto set jml_acc='"
            CMDSQL = CMDSQL + CStr(M_objrs.RecordCount) + "' where agent='"
            CMDSQL = CMDSQL + Trim(m_objrs_rpt("agent")) + "' and spvcode='"
            CMDSQL = CMDSQL + Trim(m_objrs_rpt("spvcode")) + "'"
            M_RPTCONN.execute CMDSQL
            
            Set M_objrs = Nothing
            
            m_objrs_rpt.MoveNext
        Wend
    End If
    
    Set m_objrs_rpt = Nothing
End Sub



Private Sub SHOW_PRN()
    FrmMgmReport.RPT.RetrieveDataFiles
    FrmMgmReport.RPT.WindowLeft = 0
    FrmMgmReport.RPT.WindowTop = 0
    FrmMgmReport.RPT.WindowState = crptMaximized
    FrmMgmReport.RPT.WindowShowPrintBtn = True
    FrmMgmReport.RPT.WindowShowRefreshBtn = True
    FrmMgmReport.RPT.WindowShowSearchBtn = True
    FrmMgmReport.RPT.WindowShowPrintSetupBtn = True
    FrmMgmReport.RPT.WindowControls = True
    FrmMgmReport.RPT.PrintReport
    'RPT.Action = 1
    'RPT.Reset
End Sub

Private Sub CloseWskReq(ByVal Index As Integer)

    WskRequest(Index).Close
    Unload WskRequest(Index)
    
    JmlKoneksiReq = JmlKoneksiReq - 1
    
End Sub


Private Sub WskRequest_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    JmlKoneksiReq = JmlKoneksiReq + 1
    Load WskRequest(JmlKoneksiReq)
    WskRequest(JmlKoneksiReq).Close
    
    If JmlKoneksiReq <= MaxKoneksiReq Then
        WskRequest(JmlKoneksiReq).Accept requestID
    Else
        CloseWskReq JmlKoneksiReq
    End If
End Sub



Private Sub WskRequest_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim st As String
    
    WskRequest(Index).GetData st
    TxtOnline.text = st & vbCrLf & TxtOnline.text
    TimerTandaReq.Enabled = True
End Sub




