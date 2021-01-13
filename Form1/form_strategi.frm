VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form form_strategi 
   Caption         =   "Form Strategi"
   ClientHeight    =   10305
   ClientLeft      =   495
   ClientTop       =   405
   ClientWidth     =   19605
   LinkTopic       =   "Form5"
   ScaleHeight     =   10305
   ScaleWidth      =   19605
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   10335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19695
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   29
         Left            =   16320
         TabIndex        =   30
         Top             =   8280
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   29
            Left            =   2280
            TabIndex        =   60
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   29
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   28
         Left            =   13080
         TabIndex        =   29
         Top             =   8280
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   28
            Left            =   2280
            TabIndex        =   59
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   28
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   27
         Left            =   9840
         TabIndex        =   28
         Top             =   8280
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   27
            Left            =   2280
            TabIndex        =   58
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   27
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   26
         Left            =   6600
         TabIndex        =   27
         Top             =   8280
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   26
            Left            =   2280
            TabIndex        =   57
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   26
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   25
         Left            =   3360
         TabIndex        =   26
         Top             =   8280
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   25
            Left            =   2280
            TabIndex        =   56
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   25
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   24
         Left            =   120
         TabIndex        =   25
         Top             =   8280
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   24
            Left            =   2280
            TabIndex        =   55
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   24
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   23
         Left            =   16320
         TabIndex        =   24
         Top             =   6240
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   23
            Left            =   2280
            TabIndex        =   54
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   23
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   22
         Left            =   13080
         TabIndex        =   23
         Top             =   6240
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   22
            Left            =   2280
            TabIndex        =   53
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   22
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   21
         Left            =   9840
         TabIndex        =   22
         Top             =   6240
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   21
            Left            =   2280
            TabIndex        =   52
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   21
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   20
         Left            =   6600
         TabIndex        =   21
         Top             =   6240
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   20
            Left            =   2280
            TabIndex        =   51
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   20
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   19
         Left            =   3360
         TabIndex        =   20
         Top             =   6240
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   19
            Left            =   2280
            TabIndex        =   50
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   19
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   18
         Left            =   120
         TabIndex        =   19
         Top             =   6240
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   18
            Left            =   2280
            TabIndex        =   49
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   18
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   17
         Left            =   16320
         TabIndex        =   18
         Top             =   4200
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   17
            Left            =   2280
            TabIndex        =   48
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   17
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   16
         Left            =   13080
         TabIndex        =   17
         Top             =   4200
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   16
            Left            =   2280
            TabIndex        =   47
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   16
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   15
         Left            =   9840
         TabIndex        =   16
         Top             =   4200
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   15
            Left            =   2280
            TabIndex        =   46
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   15
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   14
         Left            =   6600
         TabIndex        =   15
         Top             =   4200
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   14
            Left            =   2280
            TabIndex        =   45
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   14
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   13
         Left            =   3360
         TabIndex        =   14
         Top             =   4200
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   13
            Left            =   2280
            TabIndex        =   44
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   13
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   12
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   12
            Left            =   2280
            TabIndex        =   43
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   12
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   11
         Left            =   16320
         TabIndex        =   12
         Top             =   2160
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   11
            Left            =   2280
            TabIndex        =   42
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   11
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   10
         Left            =   13080
         TabIndex        =   11
         Top             =   2160
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   10
            Left            =   2280
            TabIndex        =   41
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   10
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   9
         Left            =   9840
         TabIndex        =   10
         Top             =   2160
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   9
            Left            =   2280
            TabIndex        =   40
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   9
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   8
         Left            =   6600
         TabIndex        =   9
         Top             =   2160
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   8
            Left            =   2400
            TabIndex        =   39
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   8
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   7
         Left            =   3360
         TabIndex        =   8
         Top             =   2160
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   7
            Left            =   2280
            TabIndex        =   38
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   7
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   6
            Left            =   2280
            TabIndex        =   37
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   5
         Left            =   16320
         TabIndex        =   6
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   5
            Left            =   2280
            TabIndex        =   36
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   5
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   4
         Left            =   13080
         TabIndex        =   5
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   4
            Left            =   2280
            TabIndex        =   35
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   4
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   3
         Left            =   9840
         TabIndex        =   4
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   34
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   3
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   2
         Left            =   6600
         TabIndex        =   3
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   2
            Left            =   2280
            TabIndex        =   33
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   2
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   1
         Left            =   3360
         TabIndex        =   2
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   32
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   1
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Action"
            Height          =   375
            Index           =   0
            Left            =   2280
            TabIndex        =   31
            Top             =   1440
            Width           =   735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1155
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "form_strategi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    open_set = Index
    get_data_set (open_set)
End Sub

Private Sub Form_Load()
    changeframecaption
    createtable
    setheader
End Sub

Private Sub changeframecaption()
    For i = 0 To 29
        g = i + 1
        Frame2(i).Caption = "Strategi " & g
    Next i
End Sub

Private Sub createtable()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sQuery = "select * from information_schema.columns  where table_name ilike 'strategi%'"
    rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic

    If rs.RecordCount = 0 Then
        qcre = "create table strategi_detail( " & vbCrLf
        qcre = qcre & "id serial not null, " & vbCrLf
        qcre = qcre & "strategi integer, " & vbCrLf
        qcre = qcre & "nm_strategi varchar, " & vbCrLf
        qcre = qcre & "sts_pop smallint default 0, " & vbCrLf
        qcre = qcre & "sts_vl smallint default 0, " & vbCrLf
        qcre = qcre & "sts_os smallint default 0, " & vbCrLf
        qcre = qcre & "sts_ptp smallint default 0, " & vbCrLf
        qcre = qcre & "sts_on smallint default 0, " & vbCrLf
        qcre = qcre & "sts_bp smallint default 0, " & vbCrLf
        qcre = qcre & "sts_co smallint default 0, " & vbCrLf
        qcre = qcre & "sts_po smallint default 0, " & vbCrLf
        qcre = qcre & "sts_pr smallint default 0, " & vbCrLf
        qcre = qcre & "balance_min numeric, " & vbCrLf
        qcre = qcre & "balance_max numeric, " & vbCrLf
        qcre = qcre & "lpd_min timestamp, " & vbCrLf
        qcre = qcre & "lpd_max timestamp, " & vbCrLf
        qcre = qcre & "wo_min timestamp, " & vbCrLf
        qcre = qcre & "wo_max timestamp, " & vbCrLf
        qcre = qcre & "create_by varchar, " & vbCrLf
        qcre = qcre & "create_date timestamp without time zone default now() " & vbCrLf
        qcre = qcre & "); " & vbCrLf & vbclrf
        
        qcre = qcre & "create table strategi_history( " & vbCrLf
        qcre = qcre & "id serial not null, " & vbCrLf
        qcre = qcre & "id_strategi integer, " & vbCrLf
        qcre = qcre & "strategi varchar, " & vbCrLf
        qcre = qcre & "run_min timestamp, " & vbCrLf
        qcre = qcre & "run_max timestamp, " & vbCrLf
        qcre = qcre & "create_by varchar, " & vbCrLf
        qcre = qcre & "create_date timestamp without time zone default now() " & vbCrLf
        qcre = qcre & "); " & vbCrLf & vbclrf
        
        qcre = qcre & "create table strategi_participan_detail( " & vbCrLf
        qcre = qcre & "id serial not null, " & vbCrLf
        qcre = qcre & "id_strategi integer, " & vbCrLf
        qcre = qcre & "strategi varchar, " & vbCrLf
        qcre = qcre & "custid varchar, " & vbCrLf
        qcre = qcre & "statuscall_bfr varchar, " & vbCrLf
        qcre = qcre & "statuscall_aft varchar, " & vbCrLf
        qcre = qcre & "agent varchar, " & vbCrLf
        qcre = qcre & "create_by varchar, " & vbCrLf
        qcre = qcre & "create_date timestamp without time zone default now() " & vbCrLf
        qcre = qcre & "); " & vbCrLf & vbclrf
        
        qcre = qcre & "create table strategi_participan( " & vbCrLf
        qcre = qcre & "id serial not null, " & vbCrLf
        qcre = qcre & "id_strategi integer, " & vbCrLf
        qcre = qcre & "strategi varchar, " & vbCrLf
        qcre = qcre & "agent varchar, " & vbCrLf
        qcre = qcre & "create_by varchar, " & vbCrLf
        qcre = qcre & "create_date timestamp without time zone default now() " & vbCrLf
        qcre = qcre & "); " & vbCrLf & vbclrf
        
        qcre = qcre & "create table strategi_run( " & vbCrLf
        qcre = qcre & "id serial not null, " & vbCrLf
        qcre = qcre & "id_strategi integer, " & vbCrLf
        qcre = qcre & "strategi varchar, " & vbCrLf
        qcre = qcre & "run_min timestamp, " & vbCrLf
        qcre = qcre & "run_max timestamp, " & vbCrLf
        qcre = qcre & "create_by varchar, " & vbCrLf
        qcre = qcre & "create_date timestamp without time zone default now() " & vbCrLf
        qcre = qcre & "); " & vbCrLf
        
        M_OBJCONN.execute qcre
    End If
End Sub

Private Sub setheader()
    For i = 0 To 29
        ListView1(i).ColumnHeaders.clear
        ListView1(i).ColumnHeaders.ADD 1, , "Strategi", 0 * 120
        ListView1(i).ColumnHeaders.ADD 2, , "Nama Strategi", 8 * 120
        ListView1(i).ColumnHeaders.ADD 3, , "Created Date", 15 * 120
        get_data_list (i)
    Next i
End Sub

Private Sub get_data_set(idx As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sQuery = "select * from strategi_history where id_strategi = " & idx & " order by id "
    rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount <> 0 Then
        With form_setstrategi
            set_strategi_value
            For a = 1 To rs.RecordCount
                Set listItem = .ListView1(1).ListItems.ADD(, , cnull(rs("id")))
                listItem.SubItems(1) = cnull(rs("id_strategi"))
                listItem.SubItems(2) = cnull(rs("strategi"))
                listItem.SubItems(3) = Format(rs("run_min"), "yyyy-mm-dd hh:nn")
                listItem.SubItems(4) = Format(rs("run_max"), "yyyy-mm-dd hh:nn")
                listItem.SubItems(5) = cnull(rs("create_by"))
                listItem.SubItems(6) = Format(rs("create_date"), "yyyy-mm-dd hh:nn:ss")
                rs.MoveNext
            Next a
            .Label4.Caption = idx
            .Show 1
        End With
    Else
        set_strategi_value
        form_setstrategi.Label4.Caption = idx
        form_setstrategi.Show 1
    End If
End Sub

Private Sub get_data_list(idx As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sQuery = "select * from strategi_detail where strategi = " & idx & " order by id "
    rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount <> 0 Then
        For a = 1 To rs.RecordCount
            Set listItem = ListView1(idx).ListItems.ADD(, , cnull(rs("strategi")))
            listItem.SubItems(1) = cnull(rs("nm_strategi"))
            listItem.SubItems(2) = cnull(rs("create_date"))
            rs.MoveNext
        Next a
    End If
End Sub

Private Sub loop_gdl()
    For S = 0 To 29
        get_data_list (S)
    Next S
End Sub

Private Sub set_strategi_value()
    form_setstrategi.ListView1(0).ColumnHeaders.clear
    form_setstrategi.ListView1(0).ColumnHeaders.ADD 1, , "Agent", 8 * 120
    form_setstrategi.ListView1(0).ColumnHeaders.ADD 2, , "Jumlah", 8 * 120
    
    form_setstrategi.ListView1(1).ColumnHeaders.clear
    form_setstrategi.ListView1(1).ColumnHeaders.ADD 1, , "id", 0 * 120
    form_setstrategi.ListView1(1).ColumnHeaders.ADD 2, , "id_strategi", 0 * 120
    form_setstrategi.ListView1(1).ColumnHeaders.ADD 3, , "Strategi", 10 * 120
    form_setstrategi.ListView1(1).ColumnHeaders.ADD 4, , "Start", 10 * 120
    form_setstrategi.ListView1(1).ColumnHeaders.ADD 5, , "Stop", 10 * 120
    form_setstrategi.ListView1(1).ColumnHeaders.ADD 6, , "Created By", 20 * 120
    form_setstrategi.ListView1(1).ColumnHeaders.ADD 7, , "Created Date", 20 * 120
End Sub
