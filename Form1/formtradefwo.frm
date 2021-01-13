VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formtradefwo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade Fresh WO"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   14415
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   14415
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Set Ranking"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   8520
      TabIndex        =   14
      Top             =   0
      Width           =   5895
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agent"
         Height          =   6015
         Left            =   2760
         TabIndex        =   73
         Top             =   240
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CommandButton Command8 
            Caption         =   "Hide"
            Height          =   375
            Left            =   2040
            TabIndex        =   104
            Top             =   5520
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   12
            Left            =   600
            TabIndex        =   83
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   13
            Left            =   600
            TabIndex        =   82
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   14
            Left            =   600
            TabIndex        =   81
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   15
            Left            =   600
            TabIndex        =   80
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   16
            Left            =   600
            TabIndex        =   79
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   17
            Left            =   600
            TabIndex        =   78
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   18
            Left            =   600
            TabIndex        =   77
            Top             =   3480
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   19
            Left            =   600
            TabIndex        =   76
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   20
            Left            =   600
            TabIndex        =   75
            Top             =   4440
            Width           =   855
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   21
            Left            =   600
            TabIndex        =   74
            Top             =   4920
            Width           =   855
         End
         Begin VB.Label a1 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   103
            Top             =   600
            Width           =   375
         End
         Begin VB.Label a2 
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   102
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label a3 
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   101
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label a4 
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label a5 
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   99
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label a6 
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   98
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label a7 
            BackStyle       =   0  'Transparent
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   97
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label a8 
            BackStyle       =   0  'Transparent
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   96
            Top             =   3960
            Width           =   375
         End
         Begin VB.Label a9 
            BackStyle       =   0  'Transparent
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   4440
            Width           =   375
         End
         Begin VB.Label a10 
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   4920
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   12
            Left            =   1560
            TabIndex        =   93
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   13
            Left            =   1560
            TabIndex        =   92
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   14
            Left            =   1560
            TabIndex        =   91
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   15
            Left            =   1560
            TabIndex        =   90
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   16
            Left            =   1560
            TabIndex        =   89
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   17
            Left            =   1560
            TabIndex        =   88
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   18
            Left            =   1560
            TabIndex        =   87
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   19
            Left            =   1560
            TabIndex        =   86
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   20
            Left            =   1560
            TabIndex        =   85
            Top             =   4440
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   375
            Index           =   21
            Left            =   1560
            TabIndex        =   84
            Top             =   4920
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Trash"
         Height          =   255
         Left            =   3360
         TabIndex        =   60
         Top             =   6120
         Visible         =   0   'False
         Width           =   735
         Begin VB.CheckBox Check3 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   71
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   70
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   69
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   68
            Top             =   1920
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check8 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   67
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   66
            Top             =   2880
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   65
            Top             =   3360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check11 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   64
            Top             =   3840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check12 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   63
            Top             =   4320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check13 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   62
            Top             =   4800
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox Check14 
            BackColor       =   &H0080FFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   0
            TabIndex        =   61
            Top             =   5280
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         Height          =   375
         Left            =   5280
         TabIndex        =   59
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   46
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   600
         TabIndex        =   41
         Top             =   5760
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   600
         TabIndex        =   40
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   600
         TabIndex        =   39
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   38
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   37
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   36
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   35
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   34
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   33
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   32
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   5280
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   600
         TabIndex        =   29
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   106
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   108
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   107
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   11
         Left            =   1440
         TabIndex        =   58
         Top             =   5760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   10
         Left            =   1440
         TabIndex        =   57
         Top             =   5280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   9
         Left            =   1440
         TabIndex        =   56
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   8
         Left            =   1440
         TabIndex        =   55
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   7
         Left            =   1440
         TabIndex        =   54
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   6
         Left            =   1440
         TabIndex        =   53
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   5
         Left            =   1440
         TabIndex        =   52
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   51
         Top             =   2400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   50
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   49
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   48
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label l12 
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label l11 
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   5280
         Width           =   375
      End
      Begin VB.Label l10 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label l9 
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label l8 
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label l7 
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label l6 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label l5 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label l4 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label l3 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.Label l1 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   15
         Top             =   6000
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Log Trade Fresh WO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   14400
      TabIndex        =   7
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5880
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5430
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   9578
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   6000
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TRADE FRESH WO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CheckBox Check15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check All"
         Height          =   255
         Left            =   4320
         TabIndex        =   110
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Auto sms"
         Height          =   495
         Left            =   5400
         TabIndex        =   109
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8640
         TabIndex        =   105
         Top             =   -120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "formtradefwo.frx":0000
         Left            =   7200
         List            =   "formtradefwo.frx":000A
         TabIndex        =   44
         Text            =   "Sistem"
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Balance"
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
         Left            =   1080
         TabIndex        =   43
         Top             =   290
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LPD"
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
         Left            =   240
         TabIndex        =   42
         Top             =   290
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "formtradefwo.frx":001E
         Left            =   3120
         List            =   "formtradefwo.frx":0028
         TabIndex        =   12
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   5880
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H000000FF&
         Caption         =   "Auto Trade"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FF8080&
         Caption         =   "Show Log Trade"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5880
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4950
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8731
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CheckBox cek_all_payment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   5880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ranking By"
         Height          =   255
         Left            =   6240
         TabIndex        =   45
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Order By :"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   5
         Top             =   6000
         Width           =   3255
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "formtradefwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loadawal As Boolean

Private Sub header()
    ListView1.ColumnHeaders.clear
    ListView1.ColumnHeaders.ADD 1, , "ID", 10 * 0
    ListView1.ColumnHeaders.ADD 2, , "Customer ID", 10 * 120
    ListView1.ColumnHeaders.ADD 3, , "To", 20 * 120
    ListView1.ColumnHeaders.ADD 4, , "Date", 20 * 120
    ListView1.ColumnHeaders.ADD 5, , "Balance", 20 * 120

    ListView2.ColumnHeaders.clear
    ListView2.ColumnHeaders.ADD 1, , "Customer ID", 20 * 120
    ListView2.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView2.ColumnHeaders.ADD 3, , "WO DATE", 8 * 120
    ListView2.ColumnHeaders.ADD 4, , "Status Account", 8 * 120
    ListView2.ColumnHeaders.ADD 5, , "Agent", 12 * 120
    ListView2.ColumnHeaders.ADD 6, , "Balance", 10 * 120
    ListView2.ColumnHeaders.ADD 7, , "L P D", 10 * 120
    'ListView2.ColumnHeaders.ADD 8, , "jarak", 10 * 0
End Sub
Private Sub getdatefwo()
    ListView2.ListItems.clear
    query = "select *,(date(now())-date(b_d)) as jarak from mgm where agent = 'TRADEFWO' order by Pay_Dt desc"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    hit = 1
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = cnull(rs("b_d"))
             listItem.SubItems(3) = cnull(rs("f_cek_new"))
             listItem.SubItems(4) = cnull(rs("agent"))
             listItem.SubItems(5) = Format(cnull(rs("curbal")), "#,#")
             listItem.SubItems(6) = Format(cnull(rs("Pay_Dt")), "yyyy-mm-dd")
             'listItem.SubItems(7) = cnull(rs("jarak"))
             If cnull(rs("jarak")) < 20 Then
                ListView2.ListItems(hit).Checked = True
                ListView2.ListItems(hit).Bold = True
                ListView2.ListItems(hit).ForeColor = vbRed
             End If
             hit = hit + 1
        rs.MoveNext
    Wend
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount
    
    ListView1.ListItems.clear
    query1 = "select a.*, b.curbal from tblfwolog a left join mgm b on a.custid = b.custid order by id desc"
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open query1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs1.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(rs1("id")))
            listItem.SubItems(1) = cnull(rs1("custid"))
            listItem.SubItems(2) = cnull(rs1("ke"))
            listItem.SubItems(3) = Format(cnull(rs1("tgl")), "yyyy-mm-dd hh:ss")
            listItem.SubItems(4) = Format(cnull(rs1("curbal")), "#,#")
        rs1.MoveNext
    Wend
    
End Sub

Private Sub Check1_Click()
    Check2.Value = 0
End Sub

Private Sub Check10_Click()
    If Check10.Value = vbChecked Then
        qr8 = "update tblranking_sistem set sts = 1 where id = 8;" & vbCrLf
        qr8 = qr8 + "update tblranking set sts = 1 where id = 8;"
        M_OBJCONN.Execute qr8
    Else
        qr8 = "update tblranking_sistem set sts = 0 where id = 8;" & vbCrLf
        qr8 = qr8 + "update tblranking set sts = 0 where id = 8;"
        M_OBJCONN.Execute qr8
    End If

End Sub

Private Sub Check11_Click()
    If Check11.Value = vbChecked Then
        qr9 = "update tblranking_sistem set sts = 1 where id = 9;" & vbCrLf
        qr9 = qr9 + "update tblranking set sts = 1 where id = 9;"
        M_OBJCONN.Execute qr9
    Else
        qr9 = "update tblranking_sistem set sts = 0 where id = 9;" & vbCrLf
        qr9 = qr9 + "update tblranking set sts = 0 where id = 9;"
        M_OBJCONN.Execute qr9
    End If
    
End Sub

Private Sub Check12_Click()
    If Check12.Value = vbChecked Then
        qr10 = "update tblranking_sistem set sts = 1 where id = 10;" & vbCrLf
        qr10 = qr10 + "update tblranking set sts = 1 where id = 10;"
        M_OBJCONN.Execute qr10
    Else
        qr10 = "update tblranking_sistem set sts = 0 where id = 10;" & vbCrLf
        qr10 = qr10 + "update tblranking set sts = 0 where id = 10;"
        M_OBJCONN.Execute qr10
    End If
End Sub

Private Sub Check13_Click()
    If Check13.Value = vbChecked Then
        qr11 = "update tblranking_sistem set sts = 1 where id = 11;" & vbCrLf
        qr11 = qr11 + "update tblranking set sts = 1 where id = 11;"
        M_OBJCONN.Execute qr11
    Else
        qr11 = "update tblranking_sistem set sts = 0 where id = 11;" & vbCrLf
        qr11 = qr11 + "update tblranking set sts = 0 where id = 11;"
        M_OBJCONN.Execute qr11
    End If
End Sub

Private Sub Check14_Click()
    If Check14.Value = vbChecked Then
        qr12 = "update tblranking_sistem set sts = 1 where id = 12;" & vbCrLf
        qr12 = qr12 + "update tblranking set sts = 1 where id = 12;"
        M_OBJCONN.Execute qr12
    Else
        qr12 = "update tblranking_sistem set sts = 0 where id = 12;" & vbCrLf
        qr12 = qr12 + "update tblranking set sts = 0 where id = 12;"
        M_OBJCONN.Execute qr12
    End If
End Sub

Private Sub check()
    qcheck = "select * from tblranking_sistem order by id"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qcheck, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    For i = 1 To rs.RecordCount
        If i = 1 Then
            If rs!STS = 1 Then
                Check3.Value = vbChecked
            Else
                Check3.Value = vbUnchecked
            End If
        End If
        If i = 2 Then
            If rs!STS = 1 Then
                Check4.Value = vbChecked
            Else
                Check4.Value = vbUnchecked
            End If
        End If
        If i = 3 Then
            If rs!STS = 1 Then
                Check5.Value = vbChecked
            Else
                Check5.Value = vbUnchecked
            End If
        End If
        If i = 4 Then
            If rs!STS = 1 Then
                Check6.Value = vbChecked
            Else
                Check6.Value = vbUnchecked
            End If
        End If
        If i = 5 Then
            If rs!STS = 1 Then
                Check7.Value = vbChecked
            Else
                Check7.Value = vbUnchecked
            End If
        End If
        If i = 6 Then
            If rs!STS = 1 Then
                Check8.Value = vbChecked
            Else
                Check8.Value = vbUnchecked
            End If
        End If
        If i = 7 Then
            If rs!STS = 1 Then
                Check9.Value = vbChecked
            Else
                Check9.Value = vbUnchecked
            End If
        End If
        If i = 8 Then
            If rs!STS = 1 Then
                Check10.Value = vbChecked
            Else
                Check10.Value = vbUnchecked
            End If
        End If
        If i = 9 Then
            If rs!STS = 1 Then
                Check11.Value = vbChecked
            Else
                Check11.Value = vbUnchecked
            End If
        End If
        If i = 10 Then
            If rs!STS = 1 Then
                Check12.Value = vbChecked
            Else
                Check12.Value = vbUnchecked
            End If
        End If
        If i = 11 Then
            If rs!STS = 1 Then
                Check13.Value = vbChecked
            Else
                Check13.Value = vbUnchecked
            End If
        End If
        If i = 12 Then
            If rs!STS = 1 Then
                Check14.Value = vbChecked
            Else
                Check14.Value = vbUnchecked
            End If
        End If
        rs.MoveNext
    Next i
    
End Sub

Private Sub Check15_Click()
    Dim r As Integer
        
    If Check15.Value = vbChecked Then
        If ListView2.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To ListView2.ListItems.Count
            ListView2.ListItems(r).Checked = True
        Next r
    Else
        For r = 1 To ListView2.ListItems.Count
            ListView2.ListItems(r).Checked = False
        Next r
    End If
End Sub

Private Sub Check2_Click()
    Check1.Value = 0
End Sub

Private Sub Check3_Click()
    If Check3.Value = vbChecked Then
        qr1 = "update tblranking_sistem set sts = 1 where id = 1;" & vbCrLf
        qr1 = qr1 + "update tblranking set sts = 1 where id = 1;"
        M_OBJCONN.Execute qr1
    Else
        qr1 = "update tblranking_sistem set sts = 0 where id = 1;" & vbCrLf
        qr1 = qr1 + "update tblranking set sts = 0 where id = 1;"
        M_OBJCONN.Execute qr1
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = vbChecked Then
        qr2 = "update tblranking_sistem set sts = 1 where id = 2;" & vbCrLf
        qr2 = qr2 + "update tblranking set sts = 1 where id = 2;"
        M_OBJCONN.Execute qr2
    Else
        qr2 = "update tblranking_sistem set sts = 0 where id = 2;" & vbCrLf
        qr2 = qr2 + "update tblranking set sts = 0 where id = 2;"
        M_OBJCONN.Execute qr2
    End If
End Sub

Private Sub Check5_Click()
    If Check5.Value = vbChecked Then
        qr3 = "update tblranking_sistem set sts = 1 where id = 3;" & vbCrLf
        qr3 = qr3 + "update tblranking set sts = 1 where id = 3;"
        M_OBJCONN.Execute qr3
    Else
        qr3 = "update tblranking_sistem set sts = 0 where id = 3;" & vbCrLf
        qr3 = qr3 + "update tblranking set sts = 0 where id = 3;"
        M_OBJCONN.Execute qr3
    End If

End Sub

Private Sub Check6_Click()
    If Check6.Value = vbChecked Then
        qr4 = "update tblranking_sistem set sts = 1 where id = 4;" & vbCrLf
        qr4 = qr4 + "update tblranking set sts = 1 where id = 4;"
        M_OBJCONN.Execute qr4
    Else
        qr4 = "update tblranking_sistem set sts = 0 where id = 4;" & vbCrLf
        qr4 = qr4 + "update tblranking set sts = 0 where id = 4;"
        M_OBJCONN.Execute qr4
    End If
End Sub

Private Sub Check7_Click()
    If Check7.Value = vbChecked Then
        qr5 = "update tblranking_sistem set sts = 1 where id = 5;" & vbCrLf
        qr5 = qr5 + "update tblranking set sts = 1 where id = 5;"
        M_OBJCONN.Execute qr5
    Else
        qr5 = "update tblranking_sistem set sts = 0 where id = 5;" & vbCrLf
        qr5 = qr5 + "update tblranking set sts = 0 where id = 5;"
        M_OBJCONN.Execute qr5
    End If
End Sub

Private Sub Check8_Click()
    If Check8.Value = vbChecked Then
        qr6 = "update tblranking_sistem set sts = 1 where id = 6;" & vbCrLf
        qr6 = qr6 + "update tblranking set sts = 1 where id = 6;"
        M_OBJCONN.Execute qr6
    Else
        qr6 = "update tblranking_sistem set sts = 0 where id = 6;" & vbCrLf
        qr6 = qr6 + "update tblranking set sts = 0 where id = 6;"
        M_OBJCONN.Execute qr6
    End If

End Sub

Private Sub Check9_Click()
    If Check9.Value = vbChecked Then
        qr7 = "update tblranking_sistem set sts = 1 where id = 7;" & vbCrLf
        qr7 = qr7 + "update tblranking set sts = 1 where id = 7;"
        M_OBJCONN.Execute qr7
    Else
        qr7 = "update tblranking_sistem set sts = 0 where id = 7;" & vbCrLf
        qr7 = qr7 + "update tblranking set sts = 0 where id = 7;"
        M_OBJCONN.Execute qr7
    End If

End Sub

Private Sub Combo1_Click()
    
If Check1.Value = 0 And Check2.Value = 0 Then
    MsgBox "Pilih Yang Ingin di Urutkan"
    Exit Sub
End If

    If Check1.Value = 1 Then
        x = "Pay_Dt"
    ElseIf Check2.Value = 1 Then
        x = "Curbal"
    End If

    ListView2.ListItems.clear
    query = "select *,(date(now())-date(b_d)) as jarak from mgm where agent = 'TRADEFWO' order by " & x & " " & Combo1.text
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    hit = 1
    If rs.RecordCount > 0 Then
        For a = 1 To rs.RecordCount
            Set listItem = ListView2.ListItems.ADD(, , cnull(rs("custid")))
                 listItem.SubItems(1) = cnull(rs("name"))
                 listItem.SubItems(2) = cnull(rs("b_d"))
                 listItem.SubItems(3) = cnull(rs("f_cek_new"))
                 listItem.SubItems(4) = cnull(rs("agent"))
                 listItem.SubItems(5) = Format(cnull(rs("curbal")), "#,#")
                 listItem.SubItems(6) = Format(cnull(rs("Pay_Dt")), "yyyy-mm-dd")
                If cnull(rs("jarak")) < 20 Then
                    ListView2.ListItems(hit).Checked = True
                    ListView2.ListItems(hit).Bold = True
                    ListView2.ListItems(hit).ForeColor = vbRed
                End If
                hit = hit + 1
            rs.MoveNext
        Next a
    End If
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount
End Sub

Private Sub getascdesc()
    a = "select * from tblascdescfwo "
    Set B = New ADODB.Recordset
    B.CursorLocation = adUseClient
    B.Open a, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If B.RecordCount = 0 Then
        c = "INSERT into tblascdescfwo values ('ASC');"
        M_OBJCONN.Execute c
        Combo1.text = "ASC"
    Else
        Combo1.text = B!Sign
    End If
End Sub

Private Sub Combo2_Click()
    If Combo2.text = "Manual" Then
        For B = 0 To 11
            t1(B).Enabled = True
            Label6(B).Visible = False
        Next B
        Command4.Visible = True
    
        q = "select * from tblranking order by 1"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        For i = 1 To 12
            a = i - 1
            t1(a).text = r!TEAM
            r.MoveNext
        Next i
    ElseIf Combo2.text = "Sistem" Then
        For B = 0 To 11
            t1(B).Enabled = False
        Next B
        
        q = "select sum(payment) as pay, b.team from tbllunas a inner join (select userid, team from usertbl) b on a.agent = b.userid "
        q = q + "where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') and team ilike 'tl%' and team in (select userid from usertbl where userid ilike 'TL%' and aktif = 0 order by 1) group by 2 order by 1 desc"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        For i = 1 To 12
            a = i - 1
            t1(a).text = r!TEAM
            Label6(a).Caption = Format(r!pay, "#,#")
            Label6(a).Visible = True
            r.MoveNext
        Next i
    
    End If
End Sub

Private Sub Command1_Click()
    Call Form_Load
End Sub

Private Sub Command10_Click()
    
    cek = 0
    For K = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For w = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(w).Checked = True Then
            CustId = ListView2.ListItems(w).text
        
            qupdate = "update mgm set autosms = 1, flagfwo = 1  where custid = '" & CustId & "'"
            M_OBJCONN.Execute qupdate
            
        End If
    Next w
    
    MsgBox "Done"

End Sub

Private Sub Command2_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If ListView2.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView2.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView2.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView2.ListItems.Count + 1
            For col = 1 To ListView2.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = "'" + ListView2.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = ListView2.ListItems(Row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(Row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_save.ShowOpen
        a = CD_save.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
zzz:
        MsgBox "No data to export", vbInformation, Me.Caption
    End If

End Sub

Private Sub Command3_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If ListView1.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView1.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView1.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView1.ListItems.Count + 1
            For col = 1 To ListView1.ColumnHeaders.Count
            If col = 1 Then
                If col <> 5 Then
                    objExcelSheet.Cells(Row, col).Value = "'" + ListView1.ListItems(Row - 1).text
                Else
                    objExcelSheet.Cells(Row, col).Value = ListView1.ListItems(Row - 1).text
                End If
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    If col <> 5 Then
                        hasil1 = "'" + ListView1.ListItems(Row - 1).SubItems(col - 1)
                    Else
                        hasil1 = ListView1.ListItems(Row - 1).SubItems(col - 1)
                    End If
                    objExcelSheet.Cells(Row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD.ShowOpen
        a = CD.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
zzz:
        MsgBox "No data to export", vbInformation, Me.Caption
    End If

End Sub

Private Sub Command4_Click()
    For F = 0 To 11
        t1(F).BackColor = vbWhite
    Next F
    
    For a = 0 To 11
        For e = 0 To 11
            If e <> a Then
                If t1(e).text = t1(a).text Then
                    t1(e).BackColor = vbRed
                    t1(a).BackColor = vbRed
                End If
            End If
        Next e
    Next a
    
    For d = 0 To 11
        If t1(d).BackColor = vbRed Then
            MsgBox "Satu TL, Satu Ranking"
            Exit Sub
        End If
    Next d
    
        qupd = ""
    
    If Combo2.text = "Manual" Then
        For i = 1 To 12
            B = i - 1
            tl = t1(B).text
            qupd = qupd + "update tblranking set team = '" + tl + "' where id = '" & i & "';" & vbCrLf
        Next i
        M_OBJCONN.Execute qupd
    
        For c = 0 To 11
            t1(c).BackColor = vbWhite
        Next c
        
        'logranking
'        insertlog = "INSERT INTO tbl_logranking_manual values ('" + t1(0).text + "','" + t1(1).text + "','" + t1(2).text + "','" + t1(3).text + "','" + t1(4).text + "','" + t1(5).text + "','" + t1(6).text + "','" + t1(7).text + "','" + t1(8).text + "','" + t1(9).text + "','" + t1(10).text + "','" + t1(11).text + "','" + MDIForm1.Text1.text + "'); "
'        M_OBJCONN.Execute insertlog
    End If
        
    MsgBox "Set Ranking Berhasil"
    
End Sub

Private Sub isiranking()
    For B = 0 To 11
        t1(B).Enabled = False
    Next B
        
    q1 = "select userid from usertbl where userid ilike 'TL%' and userid != 'TLSKIP' and aktif = 0"
    Set r1 = New ADODB.Recordset
    r1.CursorLocation = adUseClient
    r1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        q = "select sum(payment) as pay, b.team from tbllunas a inner join (select userid, team from usertbl) b on a.agent = b.userid "
        q = q + "where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') and team ilike 'tl%' and team in (select userid from usertbl where userid ilike 'TL%' and aktif = 0 order by 1) group by 2 order by 1 desc"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        For i = 1 To r1.RecordCount
            a = i - 1
            t1(a).text = r!TEAM
            Label6(a).Caption = Format(r!pay, "#,#")
            Label6(a).Visible = True
            r.MoveNext
        Next i
    
    
'    q = "select * from tblranking order by 1"
'    Set r = New ADODB.Recordset
'    r.CursorLocation = adUseClient
'    r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    For i = 1 To 12
'        a = i - 1
'        t1(a).text = r!TEAM
'        r.MoveNext
'    Next i
End Sub

Private Sub rankingsistem()
        For i = 1 To 12
            B = i - 1
            tl = t1(B).text
            qupd = qupd + "update tblranking_sistem set team = '" + tl + "' where id = '" & i & "';" & vbCrLf
        Next i
        M_OBJCONN.Execute qupd
    
        For c = 0 To 11
            t1(c).BackColor = vbWhite
        Next c
        
'        'logranking
'        insertlog = "INSERT INTO tbl_logranking_manual values ('" + t1(0).text + "','" + t1(1).text + "','" + t1(2).text + "','" + t1(3).text + "','" + t1(4).text + "','" + t1(5).text + "','" + t1(6).text + "','" + t1(7).text + "','" + t1(8).text + "','" + t1(9).text + "','" + t1(10).text + "','" + t1(11).text + "','" + MDIForm1.Text1.text + "'); "
'        M_OBJCONN.Execute insertlog
End Sub


Private Sub autotradefwoagent()
    'max rank_a dan rank
    q = "select max(rank) rank from tbltradefwoagent"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseClient
    r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    q1 = "select max(rank_a) rank_a from tbltradefwoagent"
    Set r1 = New ADODB.Recordset
    r1.CursorLocation = adUseClient
    r1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    maxrank = r!Rank
    maxrank_a = r1!rank_a
    '================================================
    
    'tradeposition
    q2 = "select * from tbltampungtradefwoagent"
    Set r2 = New ADODB.Recordset
    r2.CursorLocation = adUseClient
    r2.Open q2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    p_rank = r2!Rank
    p_rank_a = r2!rank_a
    '===================================
    
    'getcustidtotrade
    If Check1.Value = 1 Then
        x = "Pay_Dt"
    ElseIf Check2.Value = 1 Then
        x = "Curbal"
    Else
        x = "Pay_Dt"
    End If

    ListView2.ListItems.clear
    q3 = "select custid from mgm where agent = 'TRADEFWO' order by " & x & " " & Combo1.text
    Set r3 = New ADODB.Recordset
    r3.CursorLocation = adUseClient
    r3.Open q3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    '===================================
    
    'Process
    If r3.RecordCount > 0 Then
        For i = 1 To r3.RecordCount
            If p_rank = 0 Then
                p_rank = 1
                p_rank_a = 1
                    
                q4 = "select * from tbltradefwoagent where rank = " & p_rank & " and rank_a = " & p_rank_a & " ;"
                Set r4 = New ADODB.Recordset
                r4.CursorLocation = adUseClient
                r4.Open q4, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                qp = " update tbltampungtradefwoagent set rank = " & p_rank & ", rank_a = " & p_rank_a & ";"
                qp = qp & " update mgm set agent = tbltradefwoagent.agent from tbltradefwoagent where rank = " & p_rank & " and rank_a = " & p_rank_a & " and custid = '" & r3!CustId & "' ;"
                qp = qp & " insert into tblfwolog (custid,ke,tgl) values ('" & r3!CustId & "', '" & r4!agent & "', now());"
                M_OBJCONN.Execute qp
            Else
atas:
                If p_rank < maxrank Then
                    p_rank = p_rank + 1
                    
                    q4 = "select * from tbltradefwoagent where rank = " & p_rank & " and rank_a = " & p_rank_a & " ;"
                    Set r4 = New ADODB.Recordset
                    r4.CursorLocation = adUseClient
                    r4.Open q4, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    
                    
                    If r4.RecordCount > 0 Then
                    
                        If r4!agent = "TL4" Then
                            aaa = "stopdisini"
                        End If
                    
                        qp = " update tbltampungtradefwoagent set rank = " & p_rank & ", rank_a = " & p_rank_a & ";"
                        qp = qp & " update mgm set agent = tbltradefwoagent.agent from tbltradefwoagent where rank = " & p_rank & " and rank_a = " & p_rank_a & " and custid = '" & r3!CustId & "' ;"
                        qp = qp & " insert into tblfwolog (custid,ke,tgl) values ('" & r3!CustId & "', '" & r4!agent & "', now());"
                        M_OBJCONN.Execute qp
                    Else
                        GoTo atas
                    End If
                Else
                    p_rank = 1
                    
                    If p_rank_a < maxrank_a Then
                        p_rank_a = p_rank_a + 1
                    Else
                        rank_a = 1
                        p_rank_a = 1
                    End If
                    
                    q4 = "select * from tbltradefwoagent where rank = " & p_rank & " and rank_a = " & p_rank_a & " ;"
                    Set r4 = New ADODB.Recordset
                    r4.CursorLocation = adUseClient
                    r4.Open q4, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    If r4.RecordCount > 0 Then
                    
                        qp = " update tbltampungtradefwoagent set rank = " & p_rank & ", rank_a = " & p_rank_a & ";"
                        qp = qp & " update mgm set agent = tbltradefwoagent.agent from tbltradefwoagent where rank = " & p_rank & " and rank_a = " & p_rank_a & " and custid = '" & r3!CustId & "' ;"
                        qp = qp & " insert into tblfwolog (custid,ke,tgl) values ('" & r3!CustId & "', '" & r4!agent & "', now());"
                        M_OBJCONN.Execute qp
                    Else
                        GoTo atas
                    End If
'                    If p_rank = maxrank And p_rank_a = maxrank_a Then
'                        p_rank = 0
'                        p_rank_a = 0
'                    End If
                End If
            End If
        r3.MoveNext
        Next i
    End If
    '===================================
    MsgBox "Data Berhasil di Trade"
    Call Form_Load
End Sub

Private Sub Command5_Click()

    Call autotradefwoagent

'Dim sendtoatas As Integer
'    For x = 0 To 11
'        If t1(x).BackColor = vbRed Then
'            MsgBox "Tidak Boleh Ada Ranking Yang Sama"
'            Exit Sub
'        End If
'    Next x
'
'    q = "select userid from usertbl where userid ilike 'TL%' and userid != 'TLSKIP'"
'    Set r = New ADODB.Recordset
'    r.CursorLocation = adUseClient
'    r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    q1 = "select * from tblfwotl"
'    Set r1 = New ADODB.Recordset
'    r1.CursorLocation = adUseClient
'    r1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If Check1.Value = 1 Then
'        x = "Pay_Dt"
'    ElseIf Check2.Value = 1 Then
'        x = "Curbal"
'    Else
'        x = "Pay_Dt"
'    End If
'
'    ListView2.ListItems.clear
'    q2 = "select custid from mgm where agent = 'TRADEFWO' order by " & x & " " & Combo1.text
'    'q2 = "select custid from mgm where agent = 'TRADEFWO' order by Pay_Dt desc" 'curbal " & Combo1.text
'    Set r2 = New ADODB.Recordset
'    r2.CursorLocation = adUseClient
'    r2.Open q2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    a = r1!terakhir
'    If r.RecordCount <= a Then
'        B = 1
'    Else
'        B = a + 1
'    End If
'    c = r2.RecordCount
'    For e = 1 To c
'atas:
'        If Combo2.text = "Manual" Then
'            q4 = "select * from tblranking where id = '" & B & "' and team in (select userid from usertbl where userid ilike 'TL%' and aktif = 0 order by 1) "
'            Set r4 = New ADODB.Recordset
'            r4.CursorLocation = adUseClient
'            r4.Open q4, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If r4.RecordCount = 0 Then
'                GoTo lanjut
'            End If
'        ElseIf Combo2.text = "Sistem" Then
'            q4 = "select * from tblranking_sistem where id = '" & B & "' and team in (select userid from usertbl where userid ilike 'TL%' and aktif = 0 order by 1) "
'            Set r4 = New ADODB.Recordset
'            r4.CursorLocation = adUseClient
'            r4.Open q4, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If r4.RecordCount = 0 Then
'                sendtoatas = 1
'                GoTo lanjut
'            Else
'                sendtoatas = 0
'            End If
'        End If
'
'        d = r4!TEAM
'
'        q3 = "insert into tblfwolog (custid, ke, tgl) values ('" + r2!CustId + "', '" + d + "', now() );" & vbCrLf
'        q3 = q3 + "update mgm set agent = '" + d + "' where custid = '" + r2!CustId + "';" & vbCrLf
'        q3 = q3 + "update tblfwotl set terakhir = " & B & "; " & vbCrLf
'        q3 = q3 + "update tblascdescfwo set sign = '" + Combo1.text + "'; "
'        M_OBJCONN.Execute q3
'
'        r2.MoveNext
'lanjut:
'        If B = r.RecordCount Then
'            B = 1
'        Else
'            B = B + 1
'        End If
'
'        If sendtoatas = 1 Then
'            GoTo atas:
'        End If
'    Next e
'
'    MsgBox "Data Berhasil di Trade"
'    Call Form_Load
End Sub

Private Sub Command6_Click()
    'M_OBJCONN.Execute "update tblfwotl set terakhir = 0"
    M_OBJCONN.Execute "update tbltampungtradefwoagent set rank = 0, rank_a=0;"
    MsgBox "Refreshed"
    Text1.text = "0"
    Text2.text = "0"
    
End Sub

Private Sub Command7_Click()
    If formtradefwo.Width = 14505 Then
        formtradefwo.Width = 19905
    Else
        formtradefwo.Width = 14505
    End If
End Sub

Private Sub toagent(a As Integer)
    '--loopagetnsebanyak
    qs = " select max(jml) from ( " & vbCrLf
    '--looprank
    qs1 = " select count(rank) jml, rank from ( " & vbCrLf
    
    qs2 = " select * from ( " & vbCrLf
    qs2 = qs2 & " select  a.*,b.team from ( " & vbCrLf
    qs2 = qs2 & " select sum(payment) payment,agent from tbllunas where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') group by 2 " & vbCrLf
    qs2 = qs2 & " ) a left join usertbl b on a.agent = b.userid " & vbCrLf
    qs2 = qs2 & " ) a, " & vbCrLf
    qs2 = qs2 & " ( " & vbCrLf
    qs2 = qs2 & " select team,row_number() over() rank from ( " & vbCrLf
    qs2 = qs2 & " select sum(a.payment) payment,b.team from ( " & vbCrLf
    qs2 = qs2 & " select sum(payment) payment,agent from tbllunas where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') group by 2 " & vbCrLf
    qs2 = qs2 & " ) a left join usertbl b on a.agent = b.userid where team ilike 'TL%' group by 2 order by 1 desc ) a " & vbclrf
    qs2 = qs2 & " ) " & vbCrLf
    qs2 = qs2 & " b where a.team = b.team and rank = " & a & " and agent not ilike 'TL%' order by b.rank asc, payment desc " & vbCrLf
    
    qs1 = qs1 & qs2 & " ) c group by 2 " & vbclrf
    qs = qs & qs1 & " ) d " & vbclrf

    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient
    rec.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set rec1 = New ADODB.Recordset
    rec1.CursorLocation = adUseClient
    rec1.Open qs1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set rec2 = New ADODB.Recordset
    rec2.CursorLocation = adUseClient
    rec2.Open qs2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    For i = 12 To 21
        t1(i).text = ""
        Label6(i).Caption = ""
    Next i
    
    If rec2.RecordCount > 0 Then
        For i = 12 To 12 + rec1!jml - 1
            t1(i).text = rec2!agent
            Label6(i).Caption = Format(rec2!Payment, "#,#")
            rec2.MoveNext
        Next i
    End If
    
End Sub

Private Sub Command8_Click()
    Frame5.Visible = False
End Sub

Private Sub Command9_Click()
Call unionagenttl
End Sub

Private Sub Form_Load()
    loadawal = True
    Call header
    Call getdatefwo
    Call getascdesc
    Call isiranking
    Call lastrank
    Call rankingsistem
    loadawal = False
    'Call check
End Sub

Private Sub lastrank()
'    q1 = "select * from tblfwotl"
'    Set r1 = New ADODB.Recordset
'    r1.CursorLocation = adUseClient
'    r1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    q1 = "select * from tbltampungtradefwoagent"
    Set r1 = New ADODB.Recordset
    r1.CursorLocation = adUseClient
    r1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Text1.text = r1!Rank
    Text2.text = r1!rank_a
End Sub

Private Sub l1_Click()
    Frame5.Visible = True
    Call toagent(1)
End Sub

Private Sub l10_Click()
    Frame5.Visible = True
    Call toagent(10)
End Sub

Private Sub l11_Click()
    Frame5.Visible = True
    Call toagent(11)
End Sub


Private Sub unionagenttl()
    qs1 = " select count(rank) jml, rank from ( " & vbCrLf
    
    qs2 = " select * from ( " & vbCrLf
    qs2 = qs2 & " select  a.*,b.team from ( " & vbCrLf
    qs2 = qs2 & " select sum(payment) payment,agent from tbllunas where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') group by 2 " & vbCrLf
    qs2 = qs2 & " ) a left join usertbl b on a.agent = b.userid " & vbCrLf
    qs2 = qs2 & " ) a, " & vbCrLf
    qs2 = qs2 & " ( " & vbCrLf
    qs2 = qs2 & " select team,row_number() over() rank from ( " & vbCrLf
    qs2 = qs2 & " select sum(a.payment) payment,b.team from ( " & vbCrLf
    qs2 = qs2 & " select sum(payment) payment,agent from tbllunas where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') group by 2 " & vbCrLf
    qs2 = qs2 & " ) a left join usertbl b on a.agent = b.userid where team ilike 'TL%' group by 2 order by 1 desc ) a " & vbclrf
    qs2 = qs2 & " ) " & vbCrLf
    qs2 = qs2 & " b where a.team = b.team order by b.rank asc, payment desc " & vbCrLf
    
    qs1 = qs1 & qs2 & " ) c group by 2 " & vbclrf

    Set rec1 = New ADODB.Recordset
    rec1.CursorLocation = adUseClient
    rec1.Open qs1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    a = rec1.RecordCount
    
        qs2 = ""
    For i = 1 To a
        qs2 = qs2 & " select *, row_number() over() rank_a from ( " & vbCrLf
        qs2 = qs2 & " select * from ( " & vbCrLf
        qs2 = qs2 & " select  a.*,b.team from ( " & vbCrLf
        qs2 = qs2 & " select sum(payment) payment,agent from tbllunas where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') group by 2 " & vbCrLf
        qs2 = qs2 & " ) a left join usertbl b on a.agent = b.userid " & vbCrLf
        qs2 = qs2 & " ) a, " & vbCrLf
        qs2 = qs2 & " ( " & vbCrLf
        qs2 = qs2 & " select team,row_number() over() rank from ( " & vbCrLf
        qs2 = qs2 & " select sum(a.payment) payment,b.team from ( " & vbCrLf
        qs2 = qs2 & " select sum(payment) payment,agent from tbllunas where to_char(paydate, 'yyyy-mm') = to_char(now() - interval '1 month', 'yyyy-mm') group by 2 " & vbCrLf
        qs2 = qs2 & " ) a left join usertbl b on a.agent = b.userid where team ilike 'TL%' group by 2 order by 1 desc ) a " & vbclrf
        qs2 = qs2 & " ) " & vbCrLf
        qs2 = qs2 & " b where a.team = b.team and rank = " & i & " order by b.rank asc, payment desc " & vbCrLf
        qs2 = qs2 & " ) zzz " & vbCrLf
        If i <> a Then
            qs2 = qs2 & " UNION ALL " & vbCrLf
        End If
    Next i
        
    Set rec2 = New ADODB.Recordset
    rec2.CursorLocation = adUseClient
    rec2.Open qs2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
End Sub

Private Sub l12_Click()
    Frame5.Visible = True
    Call toagent(12)
End Sub

Private Sub l2_Click()
    Frame5.Visible = True
    Call toagent(2)
End Sub

Private Sub l3_Click()
    Frame5.Visible = True
    Call toagent(3)
End Sub

Private Sub l4_Click()
    Frame5.Visible = True
    Call toagent(4)
End Sub

Private Sub l5_Click()
    Frame5.Visible = True
    Call toagent(5)
End Sub

Private Sub l6_Click()
    Frame5.Visible = True
    Call toagent(6)
End Sub

Private Sub l7_Click()
    Frame5.Visible = True
    Call toagent(7)
End Sub

Private Sub l8_Click()
    Frame5.Visible = True
    Call toagent(8)
End Sub

Private Sub l9_Click()
    Frame5.Visible = True
    Call toagent(9)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub

Private Sub t1_LostFocus(Index As Integer)
   ' Command4_Click
End Sub
