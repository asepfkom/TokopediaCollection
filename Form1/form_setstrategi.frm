VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form_setstrategi 
   Caption         =   "Set Strategi"
   ClientHeight    =   7890
   ClientLeft      =   705
   ClientTop       =   825
   ClientWidth     =   14640
   LinkTopic       =   "Form5"
   ScaleHeight     =   7890
   ScaleWidth      =   14640
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.Frame Frame8 
         Caption         =   "Report"
         Height          =   4335
         Left            =   9120
         TabIndex        =   50
         Top             =   3480
         Width           =   5415
         Begin VB.CommandButton Command5 
            Caption         =   "Export"
            Height          =   375
            Left            =   4200
            TabIndex        =   52
            Top             =   3720
            Width           =   975
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3315
            Index           =   3
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   5847
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
      Begin VB.Frame Frame7 
         Caption         =   "Running"
         Height          =   3255
         Left            =   9120
         TabIndex        =   46
         Top             =   120
         Width           =   5415
         Begin VB.CommandButton Command4 
            Caption         =   "Reject"
            Height          =   375
            Left            =   4080
            TabIndex        =   49
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   2760
            TabIndex        =   48
            Top             =   2760
            Width           =   1095
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2475
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   4366
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
      Begin VB.Frame Frame6 
         Caption         =   "History"
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   5280
         Width           =   8895
         Begin MSComctlLib.ListView ListView1 
            Height          =   2115
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   3731
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
      Begin VB.Frame Frame5 
         Caption         =   "Set"
         Height          =   2295
         Left            =   4440
         TabIndex        =   9
         Top             =   2880
         Width           =   4575
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   960
            TabIndex        =   20
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Execute"
            Height          =   375
            Left            =   3360
            TabIndex        =   18
            Top             =   1680
            Width           =   975
         End
         Begin TDBDate6Ctl.TDBDate StartDate 
            Height          =   315
            Left            =   960
            TabIndex        =   14
            Top             =   840
            Width           =   1560
            _Version        =   65536
            _ExtentX        =   2752
            _ExtentY        =   556
            Calendar        =   "form_setstrategi.frx":0000
            Caption         =   "form_setstrategi.frx":0118
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "form_setstrategi.frx":0184
            Keys            =   "form_setstrategi.frx":01A2
            Spin            =   "form_setstrategi.frx":0200
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
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
            ForeColor       =   -2147483640
            Format          =   "dd/mm/yyyy"
            HighlightText   =   0
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxDate         =   2958465
            MinDate         =   -657434
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   1.12794198814265E-317
            CenturyMode     =   0
         End
         Begin TDBTime6Ctl.TDBTime StartTime 
            Height          =   315
            Left            =   2535
            TabIndex        =   15
            Top             =   840
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   556
            Caption         =   "form_setstrategi.frx":0228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "form_setstrategi.frx":0294
            Spin            =   "form_setstrategi.frx":02E4
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn"
            HighlightText   =   0
            Hour12Mode      =   1
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxTime         =   0.99999
            MidnightMode    =   0
            MinTime         =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__:__"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   1.02960316199441E-317
         End
         Begin TDBDate6Ctl.TDBDate EndDate 
            Height          =   315
            Left            =   960
            TabIndex        =   16
            Top             =   1260
            Width           =   1560
            _Version        =   65536
            _ExtentX        =   2752
            _ExtentY        =   556
            Calendar        =   "form_setstrategi.frx":030C
            Caption         =   "form_setstrategi.frx":0424
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "form_setstrategi.frx":0490
            Keys            =   "form_setstrategi.frx":04AE
            Spin            =   "form_setstrategi.frx":050C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
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
            ForeColor       =   -2147483640
            Format          =   "dd/mm/yyyy"
            HighlightText   =   0
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxDate         =   2958465
            MinDate         =   -657434
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   1.12794198814265E-317
            CenturyMode     =   0
         End
         Begin TDBTime6Ctl.TDBTime EndTime 
            Height          =   315
            Left            =   2535
            TabIndex        =   17
            Top             =   1260
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   556
            Caption         =   "form_setstrategi.frx":0534
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "form_setstrategi.frx":05A0
            Spin            =   "form_setstrategi.frx":05F0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn"
            HighlightText   =   0
            Hour12Mode      =   1
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxTime         =   0.99999
            MidnightMode    =   0
            MinTime         =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__:__"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   1.02960316199441E-317
         End
         Begin VB.Label Label3 
            Caption         =   "Strategi"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "End"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Start"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Preview Agent"
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   4215
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   3240
            TabIndex        =   44
            Text            =   "0"
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Check Agent"
            Height          =   375
            Left            =   3240
            TabIndex        =   43
            Top             =   1680
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1875
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   3307
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
         Begin VB.Label Label8 
            Caption         =   "Minimum"
            Height          =   255
            Left            =   3240
            TabIndex        =   45
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Filter"
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8895
         Begin VB.CommandButton Command1 
            Caption         =   "Preview"
            Height          =   375
            Left            =   7680
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000B&
            Caption         =   "WO"
            Height          =   735
            Index           =   3
            Left            =   3720
            TabIndex        =   5
            Top             =   1680
            Width           =   3735
            Begin VB.CheckBox Check12 
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   255
            End
            Begin TDBDate6Ctl.TDBDate TDBDate3 
               Height          =   315
               Left            =   600
               TabIndex        =   36
               Top             =   240
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               Calendar        =   "form_setstrategi.frx":0618
               Caption         =   "form_setstrategi.frx":0730
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "form_setstrategi.frx":079C
               Keys            =   "form_setstrategi.frx":07BA
               Spin            =   "form_setstrategi.frx":0818
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd/mm/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   1.12794198814265E-317
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate4 
               Height          =   315
               Left            =   2280
               TabIndex        =   37
               Top             =   240
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               Calendar        =   "form_setstrategi.frx":0840
               Caption         =   "form_setstrategi.frx":0958
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "form_setstrategi.frx":09C4
               Keys            =   "form_setstrategi.frx":09E2
               Spin            =   "form_setstrategi.frx":0A40
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd/mm/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   1.12794198814265E-317
               CenturyMode     =   0
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "To"
               Height          =   375
               Left            =   1920
               TabIndex        =   39
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000B&
            Caption         =   "LPD"
            Height          =   735
            Index           =   2
            Left            =   3720
            TabIndex        =   4
            Top             =   960
            Width           =   3735
            Begin VB.CheckBox Check11 
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   255
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Left            =   600
               TabIndex        =   34
               Top             =   240
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               Calendar        =   "form_setstrategi.frx":0A68
               Caption         =   "form_setstrategi.frx":0B80
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "form_setstrategi.frx":0BEC
               Keys            =   "form_setstrategi.frx":0C0A
               Spin            =   "form_setstrategi.frx":0C68
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd/mm/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   1.12794198814265E-317
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate2 
               Height          =   315
               Left            =   2280
               TabIndex        =   35
               Top             =   240
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               Calendar        =   "form_setstrategi.frx":0C90
               Caption         =   "form_setstrategi.frx":0DA8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "form_setstrategi.frx":0E14
               Keys            =   "form_setstrategi.frx":0E32
               Spin            =   "form_setstrategi.frx":0E90
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd/mm/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   1.12794198814265E-317
               CenturyMode     =   0
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "To"
               Height          =   375
               Left            =   1920
               TabIndex        =   38
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000B&
            Caption         =   "Balance"
            Height          =   735
            Index           =   1
            Left            =   3720
            TabIndex        =   3
            Top             =   240
            Width           =   3735
            Begin VB.CheckBox Check10 
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   255
            End
            Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
               Height          =   375
               Left            =   600
               TabIndex        =   31
               Top             =   240
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   661
               Calculator      =   "form_setstrategi.frx":0EB8
               Caption         =   "form_setstrategi.frx":0ED8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "form_setstrategi.frx":0F44
               Keys            =   "form_setstrategi.frx":0F62
               Spin            =   "form_setstrategi.frx":0FAC
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   -99999999999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   1638405
               Value           =   500000
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
               Height          =   375
               Left            =   2280
               TabIndex        =   32
               Top             =   240
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   661
               Calculator      =   "form_setstrategi.frx":0FD4
               Caption         =   "form_setstrategi.frx":0FF4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "form_setstrategi.frx":1060
               Keys            =   "form_setstrategi.frx":107E
               Spin            =   "form_setstrategi.frx":10C8
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999
               MinValue        =   -9999999999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   1992097797
               Value           =   1000000
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "To"
               Height          =   375
               Left            =   1920
               TabIndex        =   33
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000B&
            Caption         =   "Status"
            Height          =   2175
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3495
            Begin VB.CheckBox Check9 
               Caption         =   "PR"
               Height          =   375
               Left            =   2280
               TabIndex        =   30
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox Check8 
               Caption         =   "PO"
               Height          =   375
               Left            =   2280
               TabIndex        =   29
               Top             =   840
               Width           =   855
            End
            Begin VB.CheckBox Check7 
               Caption         =   "CO"
               Height          =   375
               Left            =   2280
               TabIndex        =   28
               Top             =   360
               Width           =   855
            End
            Begin VB.CheckBox Check6 
               Caption         =   "BP"
               Height          =   375
               Left            =   1200
               TabIndex        =   27
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox Check5 
               Caption         =   "ON"
               Height          =   375
               Left            =   1200
               TabIndex        =   26
               Top             =   840
               Width           =   855
            End
            Begin VB.CheckBox Check4 
               Caption         =   "PTP"
               Height          =   375
               Left            =   1200
               TabIndex        =   25
               Top             =   360
               Width           =   855
            End
            Begin VB.CheckBox Check3 
               Caption         =   "OS"
               Height          =   375
               Left            =   120
               TabIndex        =   24
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox Check2 
               Caption         =   "VL"
               Height          =   375
               Left            =   120
               TabIndex        =   23
               Top             =   840
               Width           =   855
            End
            Begin VB.CheckBox Check1 
               Caption         =   "POP"
               Height          =   375
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   855
            End
         End
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "form_setstrategi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e, F, g, h, i As Integer
Dim J, K, l, M, n, o, p, q As String
Dim unsave As Boolean
Dim c1 As Boolean

Private Sub checka()
    If Check1.Value = 1 Then
        a = 1
    Else
        a = 0
    End If
    
    If Check2.Value = 1 Then
        b = 1
    Else
        b = 0
    End If
    
    If Check3.Value = 1 Then
        c = 1
    Else
        c = 0
    End If
    
    If Check4.Value = 1 Then
        d = 1
    Else
        d = 0
    End If
    
    If Check5.Value = 1 Then
        e = 1
    Else
        e = 0
    End If
    
    If Check6.Value = 1 Then
        F = 1
    Else
        F = 0
    End If
    
    If Check7.Value = 1 Then
        g = 1
    Else
        g = 0
    End If
    
    If Check8.Value = 1 Then
        h = 1
    Else
        h = 0
    End If
    
    If Check9.Value = 1 Then
        i = 1
    Else
        i = 0
    End If
    
    If Check10.Value = 1 Then
        If TDBNumber1.Value > 0 Then
            J = "'" & TDBNumber1.Value & "'"
        Else
            J = "Null"
        End If
        
        If TDBNumber2.Value > 0 Then
            K = "'" & TDBNumber2.Value & "'"
        Else
            K = "Null"
        End If
    Else
        J = "Null"
        K = "Null"
    End If
    
    If Check11.Value = 1 Then
        If TDBDate1.Value > 0 Then
            l = "'" & Format(TDBDate1.Value, "yyyy-mm-dd") & "'"
        Else
            l = "Null"
        End If
        
        If TDBDate2.Value > 0 Then
            M = "'" & Format(TDBDate2.Value, "yyyy-mm-dd") & "'"
        Else
            M = "Null"
        End If
    Else
        l = "Null"
        M = "Null"
    End If
    
    If Check12.Value = 1 Then
        If TDBDate3.Value > 0 Then
            n = "'" & Format(TDBDate3.Value, "yyyy-mm-dd") & "'"
        Else
            n = "Null"
        End If
        
        If TDBDate3.Value > 0 Then
            o = "'" & Format(TDBDate4.Value, "yyyy-mm-dd") & "'"
        Else
            o = "Null"
        End If
    Else
        n = "Null"
        o = "Null"
    End If
End Sub

Private Sub checkb()
    If StartDate.ValueIsNull Or EndDate.ValueIsNull Then
        MsgBox "Harap isi Tanggal"
        unsave = True
        Exit Sub
    Else
        If StartTime.ValueIsNull Or EndTime.ValueIsNull Then
            MsgBox "Harap isi Jam"
            unsave = True
            Exit Sub
        Else
            
            p = Format(StartDate.Value, "yyyy-mm-dd") & " " & Format(StartTime.Value, "hh:nn")
            q = Format(EndDate.Value, "yyyy-mm-dd") & " " & Format(EndTime.Value, "hh:nn")
        End If
    End If
End Sub

Private Sub execute()
    checka
    checkb

    If unsave = False Then
        Dim rs As ADODB.Recordset
        Dim rs1 As ADODB.Recordset
        
        c1 = False
        If c1 = True Then
            running_sistem
            For i = 1 To ListView1(2).ListItems.Count
                c_id = ListView1(2).ListItems(i).text
                
                Set rs1 = New ADODB.Recordset
                rs1.CursorLocation = adUseClient
                sQuery1 = "select * from strategi_run where run_max >= '" & p & "' and id = " & c_id
                rs1.Open sQuery1, M_OBJCONN, adOpenStatic, adLockOptimistic
                
                If rs1.RecordCount > 0 Then
                    MsgBox "Start Time yand di-Set harus lebih besar dari End Time sebelumnya"
                End If
            Next i
        End If
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        sQuery = "select * from strategi_detail where strategi = '" & Label4.Caption & "' and nm_strategi = '" & Text1.text & "'"
        rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic
        
        If rs.RecordCount = 0 Then
        
            'strategi_detail
            qins = " insert into strategi_detail" & vbCrLf
            qins = qins & " (strategi,nm_strategi,sts_pop,sts_vl,sts_os,sts_ptp,sts_on,sts_bp,sts_co,sts_po,sts_pr,balance_min,balance_max,lpd_min,lpd_max,wo_min,wo_max,create_by)" & vbCrLf
            qins = qins & " Values " & vbCrLf
            qins = qins & " ('" & Label4.Caption & "','" & Text1.text & "'," & a & "," & b & "," & c & "," & d & "," & e & "," & F & "," & g & "," & h & "," & i & "," & J & "," & K & "," & l & "," & M & "," & n & "," & o & ",'" & MDIForm1.Text1.text & "');" & vbCrLf & vbCrLf
            
            'strategi_history
            qins = qins & "insert into strategi_history" & vbCrLf
            qins = qins & "(id_strategi,strategi,run_min,run_max,create_by)" & vbCrLf
            qins = qins & "Values" & vbCrLf
            qins = qins & "('" & Label4.Caption & "','" & Text1.text & "','" & p & "','" & q & "','" & MDIForm1.Text1.text & "');" & vbCrLf
            
            'strategi_participan
            'qins = qins & "delete from strategi_participan where date(create_date) < date(now());" & vbCrLf & vbCrLf
            
            qagent = ""
            For i = 1 To ListView1(0).ListItems.Count
                If ListView1(0).ListItems(i).Checked = True Then
                    qagent = qagent + "insert into strategi_participan (id_strategi,strategi,agent,create_by) values"
                    qagent = qagent + "('" & Label4.Caption & "', '" & Text1.text & "', '" & ListView1(0).ListItems(i).text & "', '" & MDIForm1.Text1.text & "');" & vbCrLf
                End If
            Next i
              
            qins = qins & qagent & vbCrLf & vbCrLf
            
            'qins = qins & "delete from strategi_run;" & vbCrLf & vbCrLf
            qins = qins & "insert into strategi_run (id_strategi,strategi,run_min,run_max,create_by) values "
            qins = qins & "('" & Label4.Caption & "','" & Text1.text & "', '" & p & "', '" & q & "', '" & MDIForm1.Text1.text & "')"
            
            M_OBJCONN.execute qins
            MsgBox "Berhasil di-set"
        Else
            MsgBox "Campaign ini sudah ada, harap bedakan nama Campaign"
        End If
    End If
End Sub

Private Sub Check10_Click()
    If Check10.Value = 1 Then
        TDBNumber1.Enabled = True
        TDBNumber2.Enabled = True
    Else
        TDBNumber1.Enabled = False
        TDBNumber2.Enabled = False
    End If
End Sub

Private Sub Check11_Click()
    If Check11.Value = 1 Then
        TDBDate1.Enabled = True
        TDBDate2.Enabled = True
    Else
        TDBDate1.Enabled = False
        TDBDate2.Enabled = False
    End If
End Sub

Private Sub Check12_Click()
    If Check12.Value = 1 Then
        TDBDate3.Enabled = True
        TDBDate4.Enabled = True
    Else
        TDBDate3.Enabled = False
        TDBDate4.Enabled = False
    End If
End Sub

Private Sub preview()
    cek = " and ("
    If a = 1 Then
        cek = cek & "f_cek_new ilike '%pop%' or"
    End If
    If b = 1 Then
        cek = cek & " f_cek_new ilike '%vl%' or"
    End If
    If c = 1 Then
        cek = cek & " f_cek_new ilike '%os%' or"
    End If
    If d = 1 Then
        cek = cek & " f_cek_new ilike '%ptp%' or"
    End If
    If e = 1 Then
        cek = cek & " f_cek_new ilike '%on%' or"
    End If
    If F = 1 Then
        cek = cek & " f_cek_new ilike '%bp%' or"
    End If
    If g = 1 Then
        cek = cek & " f_cek_new ilike '%co%' or"
    End If
    If h = 1 Then
        cek = cek & " f_cek_new ilike '%po%' or"
    End If
    If i = 1 Then
        cek = cek & " f_cek_new ilike '%pr%' or"
    End If
    If Len(cek) > 6 Then
        cek = Left(cek, Len(cek) - 3)
        cek = cek & ")" & vbCrLf
    Else
        cek = Left(cek, Len(cek) - 6)
    End If
    
    
    If Check10.Value = 1 Then
        cek = cek & " and curbal between '" & TDBNumber1.Value & "' and '" & TDBNumber2.Value & "' " & vbCrLf
    End If
    If Check11.Value = 1 Then
        cek = cek & " and pay_dt between " & Format(l, "yyyy-mm") & " and " & Format(M, "yyyy-mm") & " " & vbCrLf
    End If
    If Check12.Value = 1 Then
        cek = cek & " and b_d between " & Format(n, "yyyy-mm") & " and " & Format(o, "yyyy-mm") & " " & vbCrLf
    End If
    
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sQuery = "select agent, count(agent) as data from mgm where 1=1 " & cek & " and agent <> '' and left(agent,1) = 'D' and agent  <> 'DECEASE' group by agent order by agent "
    rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic
    
    ListView1(0).ListItems.clear
    If rs.RecordCount <> 0 Then
        For a = 1 To rs.RecordCount
            Set listItem = ListView1(0).ListItems.ADD(, , cnull(rs("agent")))
            listItem.SubItems(1) = cnull(rs("data"))
            rs.MoveNext
        Next a
    End If
End Sub

Private Sub Check13_Click()
    If Check13.Value = 1 Then
        Dim a As Integer
        Dim b As Integer
        
        
        For i = 1 To ListView1(0).ListItems.Count
            a = ListView1(0).ListItems(i).SubItems(1)
            b = Text2.text
            If a >= b Then
                ListView1(0).ListItems(i).Checked = True
            End If
        Next i
    Else
        For i = 1 To ListView1(0).ListItems.Count
            ListView1(0).ListItems(i).Checked = False
        Next i
    End If
End Sub

Private Sub Command1_Click()
    checka
    Call preview
End Sub

Private Sub Command2_Click()
    execute
End Sub

Private Sub gethst()
    nmlist = ListView1(1).SelectedItem.SubItems(2)

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sQuery = "select * from strategi_detail where strategi = " & Label4.Caption & " and nm_strategi = '" & nmlist & "';"
    rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        Check1.Value = rs!sts_pop
        Check2.Value = rs!sts_vl
        Check3.Value = rs!sts_os
        Check4.Value = rs!sts_ptp
        Check5.Value = rs!sts_on
        Check6.Value = rs!sts_bp
        Check7.Value = rs!sts_co
        Check8.Value = rs!sts_po
        Check9.Value = rs!sts_pr
        
        TDBNumber1.Value = cnull(rs!balance_min)
        TDBNumber2.Value = cnull(rs!balance_max)
        
        TDBDate1.Value = cnull(rs!lpd_min)
        TDBDate2.Value = cnull(rs!lpd_max)
        
        TDBDate3.Value = cnull(rs!wo_min)
        TDBDate4.Value = cnull(rs!wo_max)
    End If
    
End Sub

Private Sub clear()
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check8.Value = 0
    Check9.Value = 0
    Check10.Value = 0
    Check11.Value = 0
    Check12.Value = 0
    
    TDBNumber1.Value = 0
    TDBNumber2.Value = 0
    
    TDBDate1.Value = Null
    TDBDate2.Value = Null
End Sub

Private Sub Command3_Click()
    running_sistem
End Sub

Private Sub Command4_Click()
On Error GoTo bawah
    M_OBJCONN.execute "update strategi_run set run_max = run_min where id = " & ListView1(2).SelectedItem.text
    MsgBox "Reject Berhasil"
    running_sistem
    Exit Sub
bawah:
    MsgBox "Tidak ada yang direject"
End Sub

Private Sub Command5_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    If ListView1(3).ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView1(3).ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView1(3).ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView1(3).ListItems.Count + 1
            For col = 1 To ListView1(3).ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = ListView1(3).ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = "'" + ListView1(3).ListItems(Row - 1).SubItems(col - 1)
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
        MsgBox "No data to export", vbInformation, Me.Caption
    End If
End Sub

Private Sub Form_Load()
    header_running
End Sub

Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Index = 0 Then
        ListView1(0).SortKey = ColumnHeader.Index - 1
        ListView1(0).Sorted = True
    End If
End Sub

Private Sub ListView1_DblClick(Index As Integer)
    If Index = 1 Then
        clear
        gethst
        get_report
    End If
End Sub

Private Sub running_sistem()
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sQuery = "select * from strategi_run" & vbCrLf
    sQuery = sQuery & "where" & vbCrLf
    sQuery = sQuery & "id_strategi = " & Label4.Caption & "" & vbCrLf
    sQuery = sQuery & "and run_max >= now() order by id" & vbCrLf
    rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic
    
    ListView1(2).ListItems.clear
    If rs.RecordCount <> 0 Then
        For a = 1 To rs.RecordCount
            Set listItem = ListView1(2).ListItems.ADD(, , cnull(rs("id")))
            listItem.SubItems(1) = cnull(rs("strategi"))
            listItem.SubItems(2) = Format(rs("run_min"), "yyyy-mm-dd hh:nn")
            listItem.SubItems(3) = Format(rs("run_max"), "yyyy-mm-dd hh:nn")
            listItem.SubItems(4) = cnull(rs("create_by"))
            listItem.SubItems(5) = Format(rs("create_date"), "yyyy-mm-dd hh:nn:ss")
            rs.MoveNext
        Next a
    End If
End Sub

Private Sub header_running()
    ListView1(2).ColumnHeaders.clear
    ListView1(2).ColumnHeaders.ADD 1, , "id", 0 * 120
    ListView1(2).ColumnHeaders.ADD 2, , "Strategi", 8 * 120
    ListView1(2).ColumnHeaders.ADD 3, , "Start", 12 * 120
    ListView1(2).ColumnHeaders.ADD 4, , "End", 12 * 120
    ListView1(2).ColumnHeaders.ADD 5, , "Create By", 8 * 120
    ListView1(2).ColumnHeaders.ADD 6, , "Create Date", 8 * 120
End Sub

Private Sub get_report()
    Dim rs As ADODB.Recordset
    
    Call header_report
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sQuery = "select * from strategi_participan_detail" & vbCrLf
    sQuery = sQuery & "where" & vbCrLf
    sQuery = sQuery & "id_strategi = " & ListView1(1).SelectedItem.SubItems(1) & "  and" & vbCrLf
    sQuery = sQuery & "strategi = '" & ListView1(1).SelectedItem.SubItems(2) & "' and " & vbCrLf
    sQuery = sQuery & "create_date between '" & ListView1(1).SelectedItem.SubItems(3) & "' and '" & ListView1(1).SelectedItem.SubItems(4) & "'::timestamp + interval '20 min' order by 1" & vbCrLf
    rs.Open sQuery, M_OBJCONN, adOpenStatic, adLockOptimistic
    
    ListView1(3).ListItems.clear
    If rs.RecordCount <> 0 Then
        For a = 1 To rs.RecordCount
            Set listItem = ListView1(3).ListItems.ADD(, , cnull(rs("strategi")))
            listItem.SubItems(1) = cnull(rs("CustId"))
            listItem.SubItems(2) = cnull(rs("statuscall_bfr"))
            listItem.SubItems(3) = cnull(rs("statuscall_aft"))
            listItem.SubItems(4) = cnull(rs("agent"))
            listItem.SubItems(5) = Format(rs("create_date"), "yyyy-mm-dd hh:nn:ss")
            rs.MoveNext
        Next a
    End If
End Sub

Private Sub header_report()
    ListView1(3).ColumnHeaders.clear
    ListView1(3).ColumnHeaders.ADD 1, , "STRATEGI", 0 * 120
    ListView1(3).ColumnHeaders.ADD 2, , "CUSTID", 8 * 120
    ListView1(3).ColumnHeaders.ADD 3, , "STATUS AWAL", 8 * 120
    ListView1(3).ColumnHeaders.ADD 4, , "STATUS AKHIR", 8 * 120
    ListView1(3).ColumnHeaders.ADD 5, , "AGENT", 8 * 120
    ListView1(3).ColumnHeaders.ADD 6, , "TANGGAL CALL", 8 * 120
End Sub
