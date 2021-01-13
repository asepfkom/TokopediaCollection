VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmreportproductivity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Productivity"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   16830
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   16830
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   255
      Left            =   11640
      TabIndex        =   39
      Top             =   11040
      Width           =   495
      Begin VB.Frame x 
         BackColor       =   &H00C0FFFF&
         Height          =   4455
         Left            =   0
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   16335
         Begin VB.CommandButton Command6 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   12240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   840
            Width           =   1095
         End
         Begin MSComctlLib.ListView lv5 
            Height          =   3900
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   11985
            _ExtentX        =   21140
            _ExtentY        =   6879
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate10 
            Height          =   285
            Left            =   14640
            TabIndex        =   43
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":0000
            Caption         =   "frmreportproductivity.frx":0118
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":0184
            Keys            =   "frmreportproductivity.frx":01A2
            Spin            =   "frmreportproductivity.frx":0200
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate9 
            Height          =   285
            Left            =   12840
            TabIndex        =   44
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":0228
            Caption         =   "frmreportproductivity.frx":0340
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":03AC
            Keys            =   "frmreportproductivity.frx":03CA
            Spin            =   "frmreportproductivity.frx":0428
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   8
            Left            =   12240
            TabIndex        =   46
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   9
            Left            =   14400
            TabIndex        =   45
            Top             =   240
            Width           =   495
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   18653
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483648
      TabCaption(0)   =   "Report"
      TabPicture(0)   =   "frmreportproductivity.frx":0450
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CD_save"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Under 10 Second"
      TabPicture(1)   =   "frmreportproductivity.frx":046C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "Frame7"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame10 
         Caption         =   "Detail"
         Height          =   9255
         Left            =   -64080
         TabIndex        =   73
         Top             =   600
         Width           =   5655
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   3000
            TabIndex        =   84
            Top             =   7050
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   840
            TabIndex        =   82
            Top             =   7050
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   7560
            Width           =   1095
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H80000014&
            Caption         =   "Export"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   7560
            Width           =   1215
         End
         Begin MSComctlLib.ListView ListView4 
            Height          =   5940
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   10478
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate20 
            Height          =   285
            Left            =   2640
            TabIndex        =   77
            Top             =   6480
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":0488
            Caption         =   "frmreportproductivity.frx":05A0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":060C
            Keys            =   "frmreportproductivity.frx":062A
            Spin            =   "frmreportproductivity.frx":0688
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate19 
            Height          =   285
            Left            =   840
            TabIndex        =   78
            Top             =   6480
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":06B0
            Caption         =   "frmreportproductivity.frx":07C8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":0834
            Keys            =   "frmreportproductivity.frx":0852
            Spin            =   "frmreportproductivity.frx":08B0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DC"
            Height          =   255
            Index           =   21
            Left            =   2400
            TabIndex        =   83
            Top             =   7110
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Team "
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   81
            Top             =   7080
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   80
            Top             =   6480
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   18
            Left            =   2400
            TabIndex        =   79
            Top             =   6480
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "DeskColl (Hanya Perhari)"
         Height          =   9255
         Left            =   -68640
         TabIndex        =   65
         Top             =   600
         Width           =   4455
         Begin VB.CommandButton Command15 
            BackColor       =   &H80000014&
            Caption         =   "Export"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   7560
            Width           =   1215
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   7560
            Width           =   1095
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   5940
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   10478
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate18 
            Height          =   285
            Left            =   840
            TabIndex        =   69
            Top             =   7080
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":08D8
            Caption         =   "frmreportproductivity.frx":09F0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":0A5C
            Keys            =   "frmreportproductivity.frx":0A7A
            Spin            =   "frmreportproductivity.frx":0AD8
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate17 
            Height          =   285
            Left            =   840
            TabIndex        =   70
            Top             =   6480
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":0B00
            Caption         =   "frmreportproductivity.frx":0C18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":0C84
            Keys            =   "frmreportproductivity.frx":0CA2
            Spin            =   "frmreportproductivity.frx":0D00
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   17
            Left            =   720
            TabIndex        =   72
            Top             =   6840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   71
            Top             =   6480
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "TL (Hanya Perhari)"
         Height          =   9255
         Left            =   -72120
         TabIndex        =   50
         Top             =   600
         Width           =   3375
         Begin VB.CommandButton Command13 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   7560
            Width           =   1095
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H80000014&
            Caption         =   "Export"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   7560
            Width           =   1215
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   5940
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   10478
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate16 
            Height          =   285
            Left            =   840
            TabIndex        =   61
            Top             =   7080
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":0D28
            Caption         =   "frmreportproductivity.frx":0E40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":0EAC
            Keys            =   "frmreportproductivity.frx":0ECA
            Spin            =   "frmreportproductivity.frx":0F28
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate15 
            Height          =   285
            Left            =   840
            TabIndex        =   62
            Top             =   6480
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":0F50
            Caption         =   "frmreportproductivity.frx":1068
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":10D4
            Keys            =   "frmreportproductivity.frx":10F2
            Spin            =   "frmreportproductivity.frx":1150
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   64
            Top             =   6480
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   14
            Left            =   720
            TabIndex        =   63
            Top             =   6840
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Daily"
         Height          =   9255
         Left            =   -74880
         TabIndex        =   49
         Top             =   600
         Width           =   2655
         Begin VB.CommandButton Command11 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   7560
            Width           =   1095
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H80000014&
            Caption         =   "Export"
            Height          =   375
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   7560
            Width           =   1215
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   5940
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   10478
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate14 
            Height          =   285
            Left            =   840
            TabIndex        =   55
            Top             =   7080
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":1178
            Caption         =   "frmreportproductivity.frx":1290
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":12FC
            Keys            =   "frmreportproductivity.frx":131A
            Spin            =   "frmreportproductivity.frx":1378
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate13 
            Height          =   285
            Left            =   840
            TabIndex        =   56
            Top             =   6480
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":13A0
            Caption         =   "frmreportproductivity.frx":14B8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":1524
            Keys            =   "frmreportproductivity.frx":1542
            Spin            =   "frmreportproductivity.frx":15A0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   58
            Top             =   6480
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   12
            Left            =   720
            TabIndex        =   57
            Top             =   6840
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Per Agent"
         Height          =   3855
         Left            =   120
         TabIndex        =   31
         Top             =   6900
         Width           =   16335
         Begin VB.CommandButton Command9 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   12240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H80000014&
            Caption         =   "Export"
            Height          =   375
            Left            =   13560
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   840
            Width           =   1215
         End
         Begin MSComctlLib.ListView lv6 
            Height          =   3540
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   11985
            _ExtentX        =   21140
            _ExtentY        =   6244
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate12 
            Height          =   285
            Left            =   14640
            TabIndex        =   35
            Top             =   240
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":15C8
            Caption         =   "frmreportproductivity.frx":16E0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":174C
            Keys            =   "frmreportproductivity.frx":176A
            Spin            =   "frmreportproductivity.frx":17C8
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate11 
            Height          =   285
            Left            =   12840
            TabIndex        =   36
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":17F0
            Caption         =   "frmreportproductivity.frx":1908
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":1974
            Keys            =   "frmreportproductivity.frx":1992
            Spin            =   "frmreportproductivity.frx":19F0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   11
            Left            =   12240
            TabIndex        =   38
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   10
            Left            =   14400
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Per TL"
         Height          =   3255
         Left            =   120
         TabIndex        =   22
         Top             =   3540
         Width           =   16335
         Begin VB.CommandButton Command7 
            BackColor       =   &H80000014&
            Caption         =   "Export"
            Height          =   375
            Left            =   13560
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   12240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   840
            Width           =   1095
         End
         Begin MSComctlLib.ListView lv4 
            Height          =   2940
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   11985
            _ExtentX        =   21140
            _ExtentY        =   5186
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate8 
            Height          =   285
            Left            =   14640
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":1A18
            Caption         =   "frmreportproductivity.frx":1B30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":1B9C
            Keys            =   "frmreportproductivity.frx":1BBA
            Spin            =   "frmreportproductivity.frx":1C18
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate7 
            Height          =   285
            Left            =   12840
            TabIndex        =   26
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":1C40
            Caption         =   "frmreportproductivity.frx":1D58
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":1DC4
            Keys            =   "frmreportproductivity.frx":1DE2
            Spin            =   "frmreportproductivity.frx":1E40
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   6
            Left            =   14400
            TabIndex        =   28
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   7
            Left            =   12240
            TabIndex        =   27
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Day 3"
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   4140
         Width           =   16335
         Begin VB.CommandButton Command3 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   12240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   840
            Width           =   1095
         End
         Begin MSComctlLib.ListView lv3 
            Height          =   1020
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   11985
            _ExtentX        =   21140
            _ExtentY        =   1799
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate6 
            Height          =   285
            Left            =   14640
            TabIndex        =   17
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":1E68
            Caption         =   "frmreportproductivity.frx":1F80
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":1FEC
            Keys            =   "frmreportproductivity.frx":200A
            Spin            =   "frmreportproductivity.frx":2068
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate5 
            Height          =   285
            Left            =   12840
            TabIndex        =   18
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":2090
            Caption         =   "frmreportproductivity.frx":21A8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":2214
            Keys            =   "frmreportproductivity.frx":2232
            Spin            =   "frmreportproductivity.frx":2290
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   5
            Left            =   12240
            TabIndex        =   20
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   4
            Left            =   14400
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Day 2"
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   3660
         Width           =   16335
         Begin VB.CommandButton Command2 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   12240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   840
            Width           =   1095
         End
         Begin MSComctlLib.ListView lv2 
            Height          =   1020
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   11985
            _ExtentX        =   21140
            _ExtentY        =   1799
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate4 
            Height          =   285
            Left            =   14640
            TabIndex        =   10
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":22B8
            Caption         =   "frmreportproductivity.frx":23D0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":243C
            Keys            =   "frmreportproductivity.frx":245A
            Spin            =   "frmreportproductivity.frx":24B8
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate3 
            Height          =   285
            Left            =   12840
            TabIndex        =   11
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":24E0
            Caption         =   "frmreportproductivity.frx":25F8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":2664
            Keys            =   "frmreportproductivity.frx":2682
            Spin            =   "frmreportproductivity.frx":26E0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   3
            Left            =   12240
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   2
            Left            =   14400
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Report"
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   16335
         Begin VB.CommandButton Command4 
            BackColor       =   &H80000014&
            Caption         =   "Export"
            Height          =   375
            Left            =   13560
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H80000014&
            Caption         =   "Search"
            Height          =   375
            Left            =   12240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
         Begin TDBDate6Ctl.TDBDate TDBDate2 
            Height          =   285
            Left            =   14640
            TabIndex        =   2
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":2708
            Caption         =   "frmreportproductivity.frx":2820
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":288C
            Keys            =   "frmreportproductivity.frx":28AA
            Spin            =   "frmreportproductivity.frx":2908
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   285
            Left            =   12840
            TabIndex        =   3
            Top             =   240
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmreportproductivity.frx":2930
            Caption         =   "frmreportproductivity.frx":2A48
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmreportproductivity.frx":2AB4
            Keys            =   "frmreportproductivity.frx":2AD2
            Spin            =   "frmreportproductivity.frx":2B30
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin MSComctlLib.ListView lv1 
            Height          =   2340
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   11985
            _ExtentX        =   21140
            _ExtentY        =   4128
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   225
            Left            =   12240
            TabIndex        =   47
            Top             =   1200
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   397
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   12240
            TabIndex        =   48
            Top             =   1680
            Width           =   3375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   255
            Index           =   1
            Left            =   14400
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Index           =   0
            Left            =   12240
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmreportproductivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvvvvvvv(clistview As ListView, gettdbdate1 As Date, gettdbdate2 As Date)
    Dim rs As New ADODB.Recordset
    Dim list As listItem
    Dim Strsql As String

    clistview.ListItems.CLEAR
    clistview.ColumnHeaders.CLEAR
    With clistview.ColumnHeaders
        .ADD , , "DATE", 120 * 10
        .ADD , , "TEAM", 80 * 10
        .ADD , , "DC", 120 * 10
        .ADD , , "CUSTID", 120 * 10
        .ADD , , "DESTINATION", 120 * 10
    End With
    
    date1 = Format(gettdbdate1, "yyyy-mm-dd") & " " & "00:00:00"
    date2 = Format(gettdbdate2, "yyyy-mm-dd") & " " & "23:59:59"
    
    day1 = Format(gettdbdate1, "yyyy-mm-dd")
    day2 = Format(gettdbdate2, "yyyy-mm-dd")
       
    Strsql = " select '" & day1 & "'||'-'||'" & day2 & "' as tgl,usertbl.team, username, custid , destination from ( " & vbCrLf
    Strsql = Strsql + "  select username, custid, destination " & vbCrLf
    Strsql = Strsql + "     from (  " & vbCrLf
    Strsql = Strsql + "  select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null and ststelpwith in ('CH','PIC') order by id desc) a,   " & vbCrLf
    Strsql = Strsql + "  (  " & vbCrLf
    Strsql = Strsql + "  select * from (  " & vbCrLf
    Strsql = Strsql + "       SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)  " & vbCrLf
    Strsql = Strsql + "          ) a   " & vbCrLf
    Strsql = Strsql + "  ) b where a.uniqcti = b.unique_id and duration < 10 " & vbCrLf
    Strsql = Strsql + " ) a, usertbl where a.username = usertbl.userid  order by team, username"
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    zzz = rs.RecordCount
    
    
    clistview.ListItems.CLEAR
    
    While Not rs.EOF
        Set list = clistview.ListItems.ADD(, , rs!TGL)
        list.SubItems(1) = cnull(rs!TEAM)
        list.SubItems(2) = cnull(rs!UserName)
        list.SubItems(3) = cnull(rs!CustId)
        list.SubItems(4) = cnull(rs!Destination)
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

Private Sub lvvvvvv(clistview As ListView, gettdbdate1 As Date, gettdbdate2 As Date)
    Dim rs As New ADODB.Recordset
    Dim list As listItem
    Dim Strsql As String

    clistview.ListItems.CLEAR
    clistview.ColumnHeaders.CLEAR
    With clistview.ColumnHeaders
        .ADD , , "DATE", 120 * 10
        .ADD , , "TEAM", 80 * 10
        .ADD , , "DC", 120 * 10
        .ADD , , "JUMLAH", 120 * 10
    End With
    
    date1 = Format(gettdbdate1, "yyyy-mm-dd") & " " & "00:00:00"
    date2 = Format(gettdbdate2, "yyyy-mm-dd") & " " & "23:59:59"
    
    day1 = Format(gettdbdate1, "yyyy-mm-dd")
    day2 = Format(gettdbdate2, "yyyy-mm-dd")
       
    Strsql = " select '" & day1 & "'||'-'||'" & day2 & "' as tgl,usertbl.team, username, jumlah  from ( " & vbCrLf
    Strsql = Strsql + "  select username, count(username) as jumlah " & vbCrLf
    Strsql = Strsql + "     from (  " & vbCrLf
    Strsql = Strsql + "  select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null and ststelpwith in ('CH','PIC') order by id desc) a,   " & vbCrLf
    Strsql = Strsql + "  (  " & vbCrLf
    Strsql = Strsql + "  select * from (  " & vbCrLf
    Strsql = Strsql + "    " & vbCrLf
    Strsql = Strsql + "    " & vbCrLf
    Strsql = Strsql + "       SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)  " & vbCrLf
    Strsql = Strsql + "          ) a   " & vbCrLf
    Strsql = Strsql + "  ) b where a.uniqcti = b.unique_id and duration < 10 group by 1 " & vbCrLf
    Strsql = Strsql + " ) a, usertbl where a.username = usertbl.userid  order by 2"
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    zzz = rs.RecordCount
    
    
    clistview.ListItems.CLEAR
    
    While Not rs.EOF
        Set list = clistview.ListItems.ADD(, , rs!TGL)
        list.SubItems(1) = cnull(rs!TEAM)
        list.SubItems(2) = cnull(rs!UserName)
        list.SubItems(3) = cnull(rs!JUMLAH)
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

Private Sub lvvvvv(clistview As ListView, gettdbdate1 As Date, gettdbdate2 As Date)
    Dim rs As New ADODB.Recordset
    Dim list As listItem
    Dim Strsql As String

    clistview.ListItems.CLEAR
    clistview.ColumnHeaders.CLEAR
    With clistview.ColumnHeaders
        .ADD , , "DATE", 120 * 10
        .ADD , , "TEAM", 80 * 10
        .ADD , , "JUMLAH", 120 * 10
    End With
    
    date1 = Format(gettdbdate1, "yyyy-mm-dd") & " " & "00:00:00"
    date2 = Format(gettdbdate2, "yyyy-mm-dd") & " " & "23:59:59"
    
    day1 = Format(gettdbdate1, "yyyy-mm-dd")
    day2 = Format(gettdbdate2, "yyyy-mm-dd")
    
    Strsql = " select '" & day1 & "'||'-'||'" & day2 & "' as tgl, usertbl.team, count(username) jumlah from ( " & vbCrLf
    Strsql = Strsql + "  select username " & vbCrLf
    Strsql = Strsql + "     from (  " & vbCrLf
    Strsql = Strsql + "  select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null and ststelpwith in ('CH','PIC') order by id desc) a,   " & vbCrLf
    Strsql = Strsql + "  (  " & vbCrLf
    Strsql = Strsql + "  select * from (  " & vbCrLf
    Strsql = Strsql + "    " & vbCrLf
    Strsql = Strsql + "    " & vbCrLf
    Strsql = Strsql + "       SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)  " & vbCrLf
    Strsql = Strsql + "          ) a   " & vbCrLf
    Strsql = Strsql + "  ) b where a.uniqcti = b.unique_id and duration < 10 " & vbCrLf
    Strsql = Strsql + " ) a, usertbl where a.username = usertbl.userid group by 2 order by 2"

    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    zzz = rs.RecordCount
    
    
    clistview.ListItems.CLEAR
    
    While Not rs.EOF
        Set list = clistview.ListItems.ADD(, , rs!TGL)
        list.SubItems(1) = cnull(rs!TEAM)
        list.SubItems(2) = cnull(rs!JUMLAH)
        rs.MoveNext
    Wend
    Set rs = Nothing

End Sub

Private Sub lvvvv(clistview As ListView, gettdbdate1 As Date, gettdbdate2 As Date)
    Dim rs As New ADODB.Recordset
    Dim list As listItem
    Dim Strsql As String

    clistview.ListItems.CLEAR
    clistview.ColumnHeaders.CLEAR
    With clistview.ColumnHeaders
        .ADD , , "DATE", 120 * 10
        .ADD , , "JUMLAH", 120 * 10
    End With
    
    
    datejrk1 = Format(gettdbdate1, "yyyy-mm-dd")
    datejrk2 = Format(gettdbdate2, "yyyy-mm-dd")
    
    
    q = "select  '" + datejrk1 + "'::date - '" + datejrk2 + "'::date as jarak"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    jrk = -(rs!jarak)
    
    If jrk = 0 Then
        z = 1
    ElseIf jrk > 0 Then
        z = jrk + 1
        batas = z
    End If
    
    Strsql = ""
    
    For i = 1 To z
        Tanggal = datejrk1
        If i = 1 Then
            date1 = Tanggal & " 00:00:00"
            date2 = Tanggal & " 23:59:59"
        ElseIf i > 1 Then
            'Tanggal = datejrk1
            If i = 2 Then
                Tanggalu = Format(DateAdd("d", 1, Tanggal), "yyyy-mm-dd")
            Else
                Tanggalu = Format(DateAdd("d", 1, Tanggalu), "yyyy-mm-dd")
            End If
            date1 = Tanggalu & " 00:00:00"
            date2 = Tanggalu & " 23:59:59"
        End If
    
    
    Strsql = Strsql + " select " & vbCrLf
        If i = 1 Then
            Strsql = Strsql + " '" & Tanggal & "'::date as tgl, " & vbCrLf
        ElseIf i > 1 Then
            Strsql = Strsql + " '" & Tanggalu & "'::date as tgl, " & vbCrLf
        End If
    Strsql = Strsql + " count(*) as jumlah " & vbCrLf
    Strsql = Strsql + "     from (  " & vbCrLf
    Strsql = Strsql + "  select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null and ststelpwith in ('CH','PIC') order by id desc) a,   " & vbCrLf
    Strsql = Strsql + "  (  " & vbCrLf
    Strsql = Strsql + "  select * from (  " & vbCrLf
    Strsql = Strsql + "    " & vbCrLf
    Strsql = Strsql + "    " & vbCrLf
    Strsql = Strsql + "       SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)  " & vbCrLf
    Strsql = Strsql + "          ) a   " & vbCrLf
    Strsql = Strsql + "  ) b where a.uniqcti = b.unique_id and duration < 10    "
    
    If z > 1 Then
        If i < z Then
            Strsql = Strsql + " UNION ALL " & vbCrLf
        End If
    End If
    
    Next i
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    zzz = rs.RecordCount
    
    
    clistview.ListItems.CLEAR
    
    While Not rs.EOF
        Set list = clistview.ListItems.ADD(, , rs!TGL)
        list.SubItems(1) = cnull(rs!JUMLAH)
        rs.MoveNext
    Wend
    Set rs = Nothing

End Sub

Private Sub lvvv(clistview As ListView, gettdbdate1 As Date, gettdbdate2 As Date)
    Dim rs As New ADODB.Recordset
    Dim list As listItem
    Dim Strsql As String


    clistview.ColumnHeaders.CLEAR
    With clistview.ColumnHeaders
        .ADD , , "DATE", 120 * 10
        .ADD , , "LOGIN TIME", 120 * 10
        .ADD , , "USERID", 120 * 10
        .ADD , , "TEAM", 120 * 10
        .ADD , , "ATTEMPT", 120 * 10
        .ADD , , "CONNECT", 120 * 10
        .ADD , , "CONNECT DUR", 120 * 10
        .ADD , , "CONTACT", 120 * 10
        .ADD , , "CONTACT DUR", 120 * 10
        .ADD , , "NO CONTACT", 120 * 10
        .ADD , , "ACCOUNTCALLED", 120 * 10
        .ADD , , "CONTACTPERACC", 120 * 10
        .ADD , , "AVERAGE CONNECT", 120 * 10
        .ADD , , "CONNECT/ATTEMPT", 120 * 10
        .ADD , , "AVERAGE CONNECT DURATION", 120 * 10
        .ADD , , "CONTACT/CONNECT", 120 * 10
        .ADD , , "AVERAGE CONTACT DURATION", 120 * 10
        .ADD , , "WRAP", 120 * 10
        .ADD , , "IDLE", 120 * 10
    End With
    
    date1 = Format(gettdbdate1, "yyyy-mm-dd") & " 00:00:00"
    date2 = Format(gettdbdate2, "yyyy-mm-dd") & " 23:59:59"

    Strsql = " select '" & Format(gettdbdate1, "yyyy-mm-dd") & "' as tgl, jam, abcd.*,wrap,idle from ( " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + " select a.username,team,attempt,connect,connectduration,coalesce(contact,0) as contact,coalesce(contactduration,0) as contactduration,coalesce(connect-coalesce(contact,0),0) as nocontact, coalesce(accountcalled,0) accountcalled, coalesce(contactperacc,0) contactperacc, round(connect::numeric/accountcalled::numeric,3) as callperaccount, round(connect::numeric/attempt::numeric,3) as connectperattempt, round(connectduration::numeric/connect::numeric,3) as connectdurationperconnect, round(contact::numeric/connect::numeric,3) as contactperconnect, round(contactduration::numeric/contact::numeric,3) as contactdurationpercontact from ( " & vbCrLf
    Strsql = Strsql + " select username,attempt from ( " & vbCrLf
    Strsql = Strsql + " select username, count(id) as attempt from ( " & vbCrLf
    Strsql = Strsql + " SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "'' and username <> ''''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)  " & vbCrLf
    Strsql = Strsql + " ) a group by 1 order by 1  " & vbCrLf
    Strsql = Strsql + " ) a) a " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " left join " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + " select username,count(a.custid) as connect, sum(duration) as connectduration  from ( " & vbCrLf
    Strsql = Strsql + " select custid, case when disposition = 'ANSWERED' then 1 else 0 end abc, username " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " (select * from ( " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "'' and username <> ''''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)) a " & vbCrLf
    Strsql = Strsql + "          " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id " & vbCrLf
    Strsql = Strsql + " ) a,  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " (select custid, case when disposition = 'ANSWERED' then 1 else 0 end abc, duration " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " (select * from ( " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "'' and username <> ''''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)) a " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id) B " & vbCrLf
    Strsql = Strsql + " where a.custid = b.custid and a.abc = b.abc " & vbCrLf
    Strsql = Strsql + "  and a.abc = 1 group by 1 " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ) abc " & vbCrLf
    Strsql = Strsql + " ) b on a.username = b.username " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " left join  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " select a.* from ( " & vbCrLf
    Strsql = Strsql + " select count(custid) as contact, coalesce(sum(duration),0) as contactduration, username " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
    Strsql = Strsql + "         ) a left join (select userid,team from usertbl) b on a.username = b.userid  " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id " & vbCrLf
    Strsql = Strsql + " and ststelpwith in ('CH','PIC') and disposition = 'ANSWERED' group by 3 order by 3 " & vbCrLf
    Strsql = Strsql + " ) a  " & vbCrLf
    Strsql = Strsql + " ) c on a.username = c.username " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " left join  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + " select username, count(custid) accountcalled " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
    Strsql = Strsql + "         ) a " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id and disposition = 'ANSWERED' group by 1 order by 1 " & vbCrLf
    Strsql = Strsql + " ) a  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ) d on a.username = d.username " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " left join " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " select username, count(custid) contactperacc " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
    Strsql = Strsql + "         ) a  " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id and disposition = 'ANSWERED' and ststelpwith in ('CH','PIC') group by 1 " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ) e on a.username = e.username left join usertbl on a.username = usertbl.userid " & vbCrLf
    Strsql = Strsql + " ) a"
    Strsql = Strsql + " ) abcd inner join ( "
    Strsql = Strsql + " select username, round((jam/3600)::numeric,2) as jam from ( " & vbCrLf
    Strsql = Strsql + " select username,extract(hour from sum(logged_out - logged_in)) * 3600 + extract(minute from sum(logged_out - logged_in)) * 60 as jam from ( " & vbCrLf
    Strsql = Strsql + " --select username, logged_in, logged_out, logged_out - logged_in as detik from ( " & vbCrLf
    Strsql = Strsql + " SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,username, logged_in, logged_out  from public.session where (logged_in between ''" & date1 & "'' and ''" & date2 & "'') and (logged_out between ''" & date1 & "'' and ''" & date2 & "'') and username <> ''''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,username CHARACTER VARYING, logged_in timestamp without time zone, logged_out timestamp without time zone)  " & vbCrLf
    Strsql = Strsql + " --) a " & vbCrLf
    Strsql = Strsql + " ) b group by 1 order by 1 " & vbCrLf
    Strsql = Strsql + " ) c ) abcde on abcd.username = abcde.username "
    
    tanggals = Format(gettdbdate1, "yymmdd")
    tt = "  select table_name from information_schema.columns  where table_name = 'tblwrapidle_" & tanggals & "'"
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open tt, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Strsql = Strsql + " left join ( " & vbCrLf
    
    If rst.RecordCount > 0 Then
        Strsql = Strsql + " select a.team, a.username, wrap, idle from ( " & vbCrLf
        Strsql = Strsql + " select team,username,sum(detik) as wrap from (  " & vbCrLf
        Strsql = Strsql + "  select username, detik,jedah, case when jedah < 6 then 'wrap' else 'idle' end as sign from (  " & vbCrLf
        Strsql = Strsql + "  select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & tanggals & vbCrLf
        Strsql = Strsql + "  ) a ) b, usertbl where sign = 'wrap' and b.username =  usertbl.userid group by 1,2 " & vbCrLf
        Strsql = Strsql + " ) a left join ( " & vbCrLf
        Strsql = Strsql + "  select team,username,sum(detik) as idle from (  " & vbCrLf
        Strsql = Strsql + "  select username, detik,jedah, case when jedah < 6 then 'wrap' else 'idle' end as sign from (  " & vbCrLf
        Strsql = Strsql + "  select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & tanggals & vbCrLf
        Strsql = Strsql + "  ) a ) b, usertbl where sign = 'idle' and b.username =  usertbl.userid group by 1,2) b on a.username = b.username " & vbCrLf
    Else
        Strsql = Strsql + " select ''::varchar username, 'tidak ditemukan' wrap,'tidak ditemukan' idle"
    End If
    Strsql = Strsql + " ) zzz on abcd.username = zzz.username"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    zzz = rs.RecordCount
    
    
    clistview.ListItems.CLEAR
    
    While Not rs.EOF
        Set list = clistview.ListItems.ADD(, , rs!TGL)
        list.SubItems(1) = cnull(rs!JAM)
        list.SubItems(2) = cnull(rs!UserName)
        list.SubItems(3) = cnull(rs!TEAM)
        list.SubItems(4) = cnull(rs!attempt)
        list.SubItems(5) = cnull(rs!Connect)
        list.SubItems(6) = cnull(rs!connectduration) & "s"
        list.SubItems(7) = cnull(rs!contact)
        list.SubItems(8) = cnull(rs!contactduration) & "s"
        list.SubItems(9) = cnull(rs!nocontact)
        list.SubItems(10) = cnull(rs!accountcalled)
        list.SubItems(11) = cnull(rs!contactperacc)
        list.SubItems(12) = cnull(rs!callperaccount)
        list.SubItems(13) = cnull(rs!connectperattempt)
        list.SubItems(14) = cnull(rs!connectdurationperconnect) & "s"
        list.SubItems(15) = cnull(rs!contactperconnect)
        list.SubItems(16) = cnull(rs!contactdurationpercontact) & "s"
        list.SubItems(17) = cnull(rs!Wrap)
        list.SubItems(18) = cnull(rs!idle)
        
        rs.MoveNext
    Wend
    Set rs = Nothing

End Sub


Private Sub lvv(clistview As ListView, gettdbdate1 As Date, gettdbdate2 As Date)
    Dim rs As New ADODB.Recordset
    Dim list As listItem
    Dim Strsql As String


    clistview.ColumnHeaders.CLEAR
    With clistview.ColumnHeaders
        .ADD , , "DATE", 120 * 10
        .ADD , , "LOGIN TIME", 120 * 10
        .ADD , , "TEAM", 120 * 10
        .ADD , , "ATTEMPT", 120 * 10
        .ADD , , "CONNECT", 120 * 10
        .ADD , , "CONNECT DUR", 120 * 10
        .ADD , , "CONTACT", 120 * 10
        .ADD , , "CONTACT DUR", 120 * 10
        .ADD , , "NO CONTACT", 120 * 10
        .ADD , , "ACCOUNT CALLED", 120 * 10
        .ADD , , "CONTACTPERACC", 120 * 10
        .ADD , , "AVERAGE CONNECT", 120 * 10
        .ADD , , "CONNECT/ATTEMPT", 120 * 10
        .ADD , , "AVERAGE CONNECT DURATION", 120 * 10
        .ADD , , "CONTACT/CONNECT", 120 * 10
        .ADD , , "AVERAGE CONTACT DURATION", 120 * 10
        .ADD , , "WRAP", 120 * 10
        .ADD , , "IDLE", 120 * 10
    End With
    
    date1 = Format(gettdbdate1, "yyyy-mm-dd") & " 00:00:00"
    date2 = Format(gettdbdate2, "yyyy-mm-dd") & " 23:59:59"
    
    Strsql = " select '" & Format(gettdbdate1, "yyyy-mm-dd") & "' as tgl, jam,abcd.*, zzz.wrap, zzz.idle from  ( " & vbCrLf
    Strsql = Strsql + " select * from  ( " & vbCrLf
    Strsql = Strsql + " select team, attemp,connect, connectduration,coalesce(contact,0) as contact,coalesce(contactduration,0) contactduration,coalesce(connect-coalesce(contact,0),0) as nocontact,coalesce(accountcalled,0) accountcalled ,coalesce(contactperacc,0) contactperacc, round(connect::numeric/accountcalled::numeric,3) as callperaccount, round(connect::numeric/attemp::numeric,3) as connectperattempt, round(connectduration::numeric/connect::numeric,3) as connectdurationperconnect, round(contact::numeric/connect::numeric,3) as contactperconnect, round(contactduration::numeric/contact::numeric,3) as contactdurationpercontact from ( " & vbCrLf
    Strsql = Strsql + "  select count(id) attemp, case when length(team) = 3 then right(team,1) else right(team,2) end::int as hit from ( " & vbCrLf
    Strsql = Strsql + "  SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)  " & vbCrLf
    Strsql = Strsql + "   ) a left join (select userid,team from usertbl) b on a.username = b.userid where team <> '' group by 2 order by hit " & vbCrLf
    Strsql = Strsql + "   ) a  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "   left join " & vbCrLf
    Strsql = Strsql + "    " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " select *, case when length(team) = 3 then right(team,1) else right(team,2) end::int as hit from ( " & vbCrLf
    Strsql = Strsql + " select team,count(a.custid) as connect, sum(duration) as connectduration  from ( " & vbCrLf
    Strsql = Strsql + " select custid, case when disposition = 'ANSWERED' then 1 else 0 end abc, team " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " (select * from ( " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)) a " & vbCrLf
    Strsql = Strsql + " left join (select userid,team from usertbl) b on a.username = b.userid " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "          " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id " & vbCrLf
    Strsql = Strsql + " ) a,  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " (select custid, case when disposition = 'ANSWERED' then 1 else 0 end abc, duration " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " (select * from ( " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)) a " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id) B " & vbCrLf
    Strsql = Strsql + " where a.custid = b.custid and a.abc = b.abc " & vbCrLf
    Strsql = Strsql + "  and a.abc = 1 group by 1 " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ) abc order by hit " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ) b on a.hit = b.hit " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " left join  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select count(custid) as contact, coalesce(sum(duration),0) as contactduration,case when length(team) = 3 then right(team,1) else right(team,2) end::int as hit " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
    Strsql = Strsql + "         ) a left join (select userid,team from usertbl) b on a.username = b.userid  " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id " & vbCrLf
    Strsql = Strsql + " and ststelpwith in ('CH','PIC') and disposition = 'ANSWERED' group by 3 order by 3 " & vbCrLf
    Strsql = Strsql + " ) c on a.hit = c.hit " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " left join  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select case when length(team) = 3 then right(team,1) else right(team,2) end::int as hit, count(custid) accountcalled " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
    Strsql = Strsql + "         ) a left join (select userid,team from usertbl) b on a.username = b.userid " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id and disposition = 'ANSWERED' group by 1 order by 1 " & vbCrLf
    Strsql = Strsql + " ) d on a.hit = d.hit " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " left join  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + " select case when length(team) = 3 then right(team,1) else right(team,2) end::int as hit, count(custid) contactperacc " & vbCrLf
    Strsql = Strsql + "    from ( " & vbCrLf
    Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
    Strsql = Strsql + " ( " & vbCrLf
    Strsql = Strsql + " select * from ( " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "  " & vbCrLf
    Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
    Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
    Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
    Strsql = Strsql + "         ) a left join (select userid,team from usertbl) b on a.username = b.userid " & vbCrLf
    Strsql = Strsql + " ) b where a.uniqcti = b.unique_id and disposition = 'ANSWERED' and ststelpwith in ('CH','PIC') group by 1 order by 1 " & vbCrLf
    Strsql = Strsql + " ) e on a.hit = e.hit " & vbCrLf
    Strsql = Strsql + " ) abc "
    Strsql = Strsql + " ) abcd "
    Strsql = Strsql + " inner join ( "
    Strsql = Strsql + " select team, round(sum(jam)::numeric/count(team)::numeric,2) as jam from ( " & vbCrLf
    Strsql = Strsql + " select username, team, jam/3600 as jam from ( " & vbCrLf
    Strsql = Strsql + " select username,extract(hour from sum(logged_out - logged_in)) * 3600 + extract(minute from sum(logged_out - logged_in)) * 60 as jam from ( " & vbCrLf
    Strsql = Strsql + " SELECT * FROM public.dblink  " & vbCrLf
    Strsql = Strsql + "      ('demodbrnd','select id,username, logged_in, logged_out  from public.session where (logged_in between ''" & date1 & "'' and ''" & date2 & "'') and (logged_out between ''" & date1 & "'' and ''" & date2 & "'') and username <> ''''')   " & vbCrLf
    Strsql = Strsql + "          AS DATA(id INTEGER,username CHARACTER VARYING, logged_in timestamp without time zone, logged_out timestamp without time zone)  " & vbCrLf
    Strsql = Strsql + " --) a " & vbCrLf
    Strsql = Strsql + " ) b group by 1 order by 1 " & vbCrLf
    Strsql = Strsql + " ) c inner join usertbl d on c.username = d.userid " & vbCrLf
    Strsql = Strsql + " ) e group by 1 order by 1 " & vbCrLf
    Strsql = Strsql + " ) abcde on abcd.team = abcde.team" & vbCrLf
    Strsql = Strsql + " left join (" & vbCrLf
    
    tanggals = Format(gettdbdate1, "yymmdd")
        
        tt = "  select table_name from information_schema.columns  where table_name = 'tblwrapidle_" & tanggals & "'"
        Set rst = New ADODB.Recordset
        rst.CursorLocation = adUseClient
        rst.Open tt, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rst.RecordCount > 0 Then
        Strsql = Strsql + " select a.team,wrap,idle from ( " & vbCrLf
        Strsql = Strsql + " select team, sum(wrap) wrap from ( " & vbCrLf
        Strsql = Strsql + " select usertbl.team,a.* from ( " & vbCrLf
        Strsql = Strsql + " select username, sum(detik) as wrap from ( " & vbCrLf
        Strsql = Strsql + " select * from ( " & vbCrLf
        Strsql = Strsql + " select username, detik,jedah, case when jedah < 6 then 'wrap' else 'idle' end as sign from ( " & vbCrLf
        Strsql = Strsql + " select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & tanggals & vbCrLf
        Strsql = Strsql + " ) a ) b where sign = 'wrap' " & vbCrLf
        Strsql = Strsql + " ) a group by 1 ) a, usertbl where a.username = usertbl.userid  " & vbCrLf
        Strsql = Strsql + " ) a group by 1) a , " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + " select * from ( " & vbCrLf
        Strsql = Strsql + " select team, sum(idle) idle from ( " & vbCrLf
        Strsql = Strsql + " select usertbl.team,a.* from ( " & vbCrLf
        Strsql = Strsql + " select username, sum(detik) as idle from ( " & vbCrLf
        Strsql = Strsql + " select * from ( " & vbCrLf
        Strsql = Strsql + " select username, detik,jedah, case when jedah < 6 then 'wrap' else 'idle' end as sign from ( " & vbCrLf
        Strsql = Strsql + " select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & tanggals & vbCrLf
        Strsql = Strsql + " ) a ) b where sign = 'idle' " & vbCrLf
        Strsql = Strsql + " ) a group by 1 ) a, usertbl where a.username = usertbl.userid  " & vbCrLf
        Strsql = Strsql + " ) a group by 1) a ) b where a.team = b.team " & vbCrLf
    Else
        Strsql = Strsql + " select ''::varchar team, 'tidak ditemukan' wrap,'tidak ditemukan' idle"
    End If
        Strsql = Strsql + " ) zzz" & vbCrLf
        Strsql = Strsql + " on abcde.team = zzz.team"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    zzz = rs.RecordCount
    
    
    clistview.ListItems.CLEAR
    
    While Not rs.EOF
        Set list = clistview.ListItems.ADD(, , rs!TGL)
        list.SubItems(1) = cnull(rs!JAM)
        list.SubItems(2) = cnull(rs!TEAM)
        list.SubItems(3) = cnull(rs!attemp)
        list.SubItems(4) = cnull(rs!Connect)
        list.SubItems(5) = cnull(rs!connectduration) & "s"
        list.SubItems(6) = cnull(rs!contact)
        list.SubItems(7) = cnull(rs!contactduration) & "s"
        list.SubItems(8) = cnull(rs!nocontact)
        list.SubItems(9) = cnull(rs!accountcalled)
        list.SubItems(10) = cnull(rs!contactperacc)
        list.SubItems(11) = cnull(rs!callperaccount)
        list.SubItems(12) = cnull(rs!connectperattempt)
        list.SubItems(13) = cnull(rs!connectdurationperconnect) & "s"
        list.SubItems(14) = cnull(rs!contactperconnect)
        list.SubItems(15) = cnull(rs!contactdurationpercontact) & "s"
        list.SubItems(16) = cnull(rs!Wrap)
        list.SubItems(17) = cnull(rs!idle)
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

Private Sub lv(clistview As ListView, gettdbdate1 As Date, gettdbdate2 As Date)
    Dim rs As New ADODB.Recordset
    Dim list As listItem
    Dim Strsql As String


    clistview.ListItems.CLEAR
    clistview.ColumnHeaders.CLEAR
    With clistview.ColumnHeaders
        .ADD , , "DATE", 120 * 10
        .ADD , , "LOGIN TIME", 120 * 10
        .ADD , , "ATTEMPT", 120 * 10
        .ADD , , "CONNECT", 120 * 10
        .ADD , , "CONNECT DUR", 120 * 10
        .ADD , , "CONTACT", 120 * 10
        .ADD , , "CONTACT DUR", 120 * 10
        .ADD , , "NO CONTACT", 120 * 10
        .ADD , , "ACCOUNT CALLED", 120 * 10
        '.ADD , , "CALL PER ACCOUNT", 120 * 10
        .ADD , , "CONTACTPERACC", 120 * 10
        .ADD , , "AVERAGE CONNECT", 120 * 10
        .ADD , , "CONNECT/ATTEMPT", 120 * 10
        .ADD , , "AVERAGE CONNECT DURATION", 120 * 10
        .ADD , , "CONTACT/CONNECT", 120 * 10
        .ADD , , "AVERAGE CONTACT DURATION", 120 * 10
        .ADD , , "WRAP", 120 * 10
        .ADD , , "IDLE", 120 * 10
    End With
    
    datejrk1 = Format(gettdbdate1, "yyyy-mm-dd") '& " 00:00:00"
    datejrk2 = Format(gettdbdate2, "yyyy-mm-dd") '& " 23:59:59"

    q = "select  '" + datejrk1 + "'::date - '" + datejrk2 + "'::date as jarak"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    jrk = -(rs!jarak)
    
    If jrk = 0 Then
        z = 1
    ElseIf jrk > 0 Then
        z = jrk + 1
        batas = z - 1
    End If
    
    ProgressBar1.Max = z
    
    Strsql = ""
    
    Label2.Caption = "Harap Tunggu"

    For i = 1 To z
    Tanggal = datejrk1
'        date1 = date1 & "00:00:00"
'        date2 = date1 & "23:59:59"
        
        If i = 1 Then
            date1 = Tanggal & " 00:00:00"
            date2 = Tanggal & " 23:59:59"
        ElseIf i > 1 Then
            'Tanggal = datejrk1
            If i = 2 Then
                Tanggalu = Format(DateAdd("d", 1, Tanggal), "yyyy-mm-dd")
            Else
                Tanggalu = Format(DateAdd("d", 1, Tanggalu), "yyyy-mm-dd")
            End If
            date1 = Tanggalu & " 00:00:00"
            date2 = Tanggalu & " 23:59:59"
        End If
        
        'cek dulu kosong apa gak
        
        CC = ""
        
        CC = CC + "  select * from ("
        CC = CC + "  SELECT count(*) attempt FROM public.dblink"
        CC = CC + "  ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "'' and username <> ''''')"
        CC = CC + "  AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) ) a, ("
        CC = CC + "    select count(custid) as contact, coalesce(sum(duration),0) as contactduration"
        CC = CC + "  from ("
        CC = CC + "  select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,"
        CC = CC + "  ("
        CC = CC + "  select * from ("
        CC = CC + "   SELECT * FROM public.dblink"
        CC = CC + "  ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')"
        CC = CC + "   AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)"
        CC = CC + "   ) a"
        CC = CC + "  ) b where a.uniqcti = b.unique_id"
        CC = CC + "  and ststelpwith in ('CH','PIC') and disposition = 'ANSWERED'  ) b"
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open CC, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        '-----------------------
        
                

        Strsql = Strsql + " select "
        If i = 1 Then
            Strsql = Strsql + " '" & Tanggal & "'::date as tgl, " & vbCrLf
        ElseIf i > 1 Then
            Strsql = Strsql + " '" & Tanggalu & "'::date as tgl, " & vbCrLf
        End If
        Strsql = Strsql + " jam,attempt,connect,connectduration,contact,contactduration,connect-contact as nocontact,accountcalled,contactperacc " & vbCrLf
        If rs!attempt <> 0 Then
            Strsql = Strsql + " ,round(connect::numeric/accountcalled::numeric,3) as callperacc, round(connect/attempt::numeric,3) as connectperattempt, round(connectduration::numeric/connect::numeric,3) as conndurationperconnect, round(contact/connect::numeric,3) as contactperconnect, " & vbCrLf
            If rs!contact <> 0 Then
                Strsql = Strsql + " round(contactduration::numeric/contact::numeric,3) as contdurpercont " & vbCrLf
            Else
                Strsql = Strsql + " '0' as contdurpercont " & vbCrLf
            End If
        Else
            Strsql = Strsql + " ,'0' as callperacc, '0' as connectperattempt, '0' as conndurationperconnect, '0' as contactperconnect, '0' as contdurpercont " & vbCrLf
        End If
        Strsql = Strsql + " , wrap, idle" & vbCrLf
        Strsql = Strsql + " From " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + " Select * From " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " SELECT count(*) attempt FROM public.dblink " & vbCrLf
        Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "'' and username <> ''''')  " & vbCrLf
        Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " )z, " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " select count(a.custid) as connect, sum(duration) as connectduration  from ( " & vbCrLf
        Strsql = Strsql + " select custid, case when disposition = 'ANSWERED' then 1 else 0 end abc " & vbCrLf
        Strsql = Strsql + "    from ( " & vbCrLf
        Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
        Strsql = Strsql + " (select * from ( " & vbCrLf
        Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
        Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
        Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)) a " & vbCrLf
        Strsql = Strsql + " ) b where a.uniqcti = b.unique_id " & vbCrLf
        Strsql = Strsql + " ) a,  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " (select custid, case when disposition = 'ANSWERED' then 1 else 0 end abc, duration " & vbCrLf
        Strsql = Strsql + "    from ( " & vbCrLf
        Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
        Strsql = Strsql + " (select * from ( " & vbCrLf
        Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
        Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
        Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING)) a " & vbCrLf
        Strsql = Strsql + " ) b where a.uniqcti = b.unique_id) B " & vbCrLf
        Strsql = Strsql + " where a.custid = b.custid and a.abc = b.abc " & vbCrLf
        Strsql = Strsql + "  and a.abc = 1 " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " )a, " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " select count(custid) as contact, coalesce(sum(duration),0) as contactduration " & vbCrLf
        Strsql = Strsql + "    from ( " & vbCrLf
        Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + " select * from ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
        Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
        Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
        Strsql = Strsql + "         ) a  " & vbCrLf
        Strsql = Strsql + " ) b where a.uniqcti = b.unique_id " & vbCrLf
        Strsql = Strsql + " and ststelpwith in ('CH','PIC') and disposition = 'ANSWERED' " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " )b, " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " select sum(call) accountcalled from ( " & vbCrLf
        Strsql = Strsql + " select custid, count(custid) call " & vbCrLf
        Strsql = Strsql + "    from ( " & vbCrLf
        Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + " select * from ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
        Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
        Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
        Strsql = Strsql + "         ) a  " & vbCrLf
        Strsql = Strsql + " ) b where a.uniqcti = b.unique_id and disposition = 'ANSWERED' group by 1 " & vbCrLf
        Strsql = Strsql + " ) a " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " )c, " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " select coalesce(sum(call),0) contactperacc   from ( " & vbCrLf
        Strsql = Strsql + " select custid, count(custid) call " & vbCrLf
        Strsql = Strsql + "    from ( " & vbCrLf
        Strsql = Strsql + " select * from mgm_hst where uniqcti is not null and uniqcti != '' and ststelpwith <> '' and ststelpwith is not null order by id desc) a,  " & vbCrLf
        Strsql = Strsql + " ( " & vbCrLf
        Strsql = Strsql + " select * from ( " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + "      SELECT * FROM public.dblink " & vbCrLf
        Strsql = Strsql + "     ('demodbrnd','select id,account_code,source,destination,destination_context,caller_id,channel,destination_channel,last_application,last_data,start_time,answer_time,end_time,   duration,billable_seconds,disposition,ama_flags,unique_id,user_field,hangup_cause,username,customer_id,campaign  from public.call_log where start_time between ''" & date1 & "'' and ''" & date2 & "''')  " & vbCrLf
        Strsql = Strsql + "         AS DATA(id INTEGER,account_code CHARACTER VARYING, source CHARACTER VARYING,destination CHARACTER VARYING, destination_context CHARACTER VARYING, caller_id CHARACTER VARYING, channel CHARACTER VARYING, destination_channel CHARACTER VARYING, last_application CHARACTER VARYING, last_data CHARACTER VARYING, start_time timestamp without time zone, answer_time timestamp without time zone, end_time timestamp without time zone, duration integer, billable_seconds integer, disposition CHARACTER VARYING, ama_flags CHARACTER VARYING, unique_id CHARACTER VARYING, user_field CHARACTER VARYING, hangup_cause CHARACTER VARYING, username CHARACTER VARYING, customer_id CHARACTER VARYING, campaign CHARACTER VARYING) " & vbCrLf
        Strsql = Strsql + "         ) a  " & vbCrLf
        Strsql = Strsql + " ) b where a.uniqcti = b.unique_id and disposition = 'ANSWERED' and ststelpwith in ('CH','PIC') group by 1 " & vbCrLf
        Strsql = Strsql + " ) a " & vbCrLf
        Strsql = Strsql + "  " & vbCrLf
        Strsql = Strsql + " ) d " & vbCrLf
        Strsql = Strsql + " ) abc " & vbCrLf
        Strsql = Strsql + " , ( " & vbCrLf
        Strsql = Strsql + " select round(sum(jam/3600)::numeric/count(username),2) as jam from ( " & vbCrLf
        Strsql = Strsql + " select username,extract(hour from sum(logged_out - logged_in)) * 3600 + extract(minute from sum(logged_out - logged_in)) * 60 as jam from ( " & vbCrLf
        Strsql = Strsql + " SELECT * FROM public.dblink  " & vbCrLf
        Strsql = Strsql + "      ('demodbrnd','select id,username, logged_in, logged_out  from public.session where (logged_in between ''" & date1 & "'' and ''" & date2 & "'') and (logged_out between ''" & date1 & "'' and ''" & date2 & "'') and username <> ''''')   " & vbCrLf
        Strsql = Strsql + "          AS DATA(id INTEGER,username CHARACTER VARYING, logged_in timestamp without time zone, logged_out timestamp without time zone)  " & vbCrLf
        Strsql = Strsql + " ) b group by 1 order by 1 " & vbCrLf
        Strsql = Strsql + " ) c " & vbCrLf
        Strsql = Strsql + " ) e "
        
        tt = ""
        
        If i = 1 Then
            tt = tt + "  select table_name from information_schema.columns  where table_name = 'tblwrapidle_" & Format(Tanggal, "yymmdd") & "'"
        ElseIf i > 1 Then
            tt = tt + "  select table_name from information_schema.columns  where table_name = 'tblwrapidle_" & Format(Tanggalu, "yymmdd") & "'"
        End If
        Set rst = New ADODB.Recordset
        rst.CursorLocation = adUseClient
        rst.Open tt, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If rst.RecordCount > 0 Then
            Strsql = Strsql + " , (select sum(detik)::varchar wrap from (" & vbCrLf
            Strsql = Strsql + " select username, detik,jedah, case when jedah < 6 then 'wrap' else 'idle' end as sign from (" & vbCrLf
            If i = 1 Then
                Strsql = Strsql + " select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & Format(Tanggal, "yymmdd") & vbCrLf
            ElseIf i > 1 Then
                Strsql = Strsql + " select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & Format(Tanggalu, "yymmdd") & vbCrLf
            End If
            Strsql = Strsql + " ) a" & vbCrLf
            Strsql = Strsql + " ) b where sign = 'wrap') f," & vbCrLf
            Strsql = Strsql + " (select sum(detik)::varchar idle from (" & vbCrLf
            Strsql = Strsql + " select username, detik,jedah, case when jedah < 6 then 'wrap' else 'idle' end as sign from (" & vbCrLf
            If i = 1 Then
                Strsql = Strsql + " select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & Format(Tanggal, "yymmdd") & vbCrLf
            ElseIf i > 1 Then
                Strsql = Strsql + " select username,(waktuyangingindikurangi - waktuyangmengurangi) as detik,round((waktuyangingindikurangi - waktuyangmengurangi)/60) as jedah from tblwrapidle_" & Format(Tanggalu, "yymmdd") & vbCrLf
            End If
            Strsql = Strsql + " ) a" & vbCrLf
            Strsql = Strsql + " ) b where sign = 'idle') g" & vbCrLf
        Else
            Strsql = Strsql + ", (select 'Tidak Ditemukan'::varchar Wrap) f, (select 'Tidak ditemukan'::varchar Idle) g"
        End If
        
        
            If z > 0 Then
                If i < batas + 1 Then
                    Strsql = Strsql + " UNION ALL " & vbCrLf
                End If
            End If
        ProgressBar1.Value = i
    Next i
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    zzz = rs.RecordCount
    
    
    clistview.ListItems.CLEAR
    
    While Not rs.EOF
        Set list = clistview.ListItems.ADD(, , rs!TGL)
        list.SubItems(1) = cnull(rs!JAM)
        list.SubItems(2) = cnull(rs!attempt)
        list.SubItems(3) = cnull(rs!Connect)
        list.SubItems(4) = cnull(rs!connectduration) & "s"
        list.SubItems(5) = cnull(rs!contact)
        list.SubItems(6) = cnull(rs!contactduration) & "s"
        list.SubItems(7) = cnull(rs!nocontact)
        list.SubItems(8) = cnull(rs!accountcalled)
        list.SubItems(9) = cnull(rs!contactperacc)
        If rs!attempt <> 0 Then
            list.SubItems(10) = cnull(rs!callperacc)
            list.SubItems(11) = cnull(rs!connectperattempt)
            list.SubItems(12) = cnull(rs!conndurationperconnect) & "s"
            list.SubItems(13) = cnull(rs!contactperconnect)
            list.SubItems(14) = cnull(rs!contdurpercont) & "s"
        End If
        list.SubItems(15) = cnull(rs!Wrap)
        list.SubItems(16) = cnull(rs!idle)
        rs.MoveNext
    Wend
    
    Label2.Caption = "Proses pencarian selesai"
      
    Set rs = Nothing
End Sub

Private Sub Command1_Click()
   Call lv(lv1, TDBDate1.Value, TDBDate2.Value)
End Sub

Private Sub Command10_Click()
    Call export(ListView1)
End Sub

Private Sub export(clistview As ListView)
        Dim objExcel As New Excel.Application
        Dim objExcelSheet As Excel.Worksheet
        Dim col, Row As Integer
        Dim a As String
        'On Error GoTo zzz
        If clistview.ListItems.Count > 0 Then
            objExcel.Workbooks.ADD
            Set objExcelSheet = objExcel.Worksheets.ADD
         
        
            For col = 1 To clistview.ColumnHeaders.Count
                objExcelSheet.Cells(1, col).Value = clistview.ColumnHeaders(col)
            Next
         
            For Row = 2 To clistview.ListItems.Count + 1
                For col = 1 To clistview.ColumnHeaders.Count
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = "'" + clistview.ListItems(Row - 1).text
                    Else
                        '" 'cararandy 29032016 "
                        Dim hasil1 As String
                            hasil1 = clistview.ListItems(Row - 1).SubItems(col - 1)
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
'zzz:
            MsgBox "No data to export", vbInformation, Me.Caption
        End If

End Sub

Private Sub Command11_Click()
    Call lvvvv(ListView1, TDBDate13.Value, TDBDate14.Value)
End Sub

Private Sub Command12_Click()
    Call export(ListView2)
End Sub

Private Sub Command13_Click()
    Call lvvvvv(ListView2, TDBDate15.Value, TDBDate15.Value)
End Sub

Private Sub Command14_Click()
    Call lvvvvvv(ListView3, TDBDate17.Value, TDBDate17.Value)
End Sub

Private Sub Command15_Click()
    Call export(ListView3)
End Sub

Private Sub Command16_Click()
    Call export(ListView4)
End Sub

Private Sub Command17_Click()
    Call lvvvvvvv(ListView4, TDBDate19, TDBDate19)
End Sub

Private Sub Command2_Click()
    Call lv(lv2, TDBDate3.Value, TDBDate4.Value)
End Sub

Private Sub Command3_Click()
    Call lv(lv3, TDBDate5.Value, TDBDate6.Value)
End Sub

Private Sub Command4_Click()
        Dim objExcel As New Excel.Application
        Dim objExcelSheet As Excel.Worksheet
        Dim col, Row As Integer
        Dim a As String
        'On Error GoTo zzz
        If lv1.ListItems.Count > 0 Then
            objExcel.Workbooks.ADD
            Set objExcelSheet = objExcel.Worksheets.ADD
         
        
            For col = 1 To lv1.ColumnHeaders.Count
                objExcelSheet.Cells(1, col).Value = lv1.ColumnHeaders(col)
            Next
         
            For Row = 2 To lv1.ListItems.Count + 1
                For col = 1 To lv1.ColumnHeaders.Count
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = "'" + lv1.ListItems(Row - 1).text
                    Else
                        '" 'cararandy 29032016 "
                        Dim hasil1 As String
                            hasil1 = lv1.ListItems(Row - 1).SubItems(col - 1)
                            objExcelSheet.Cells(Row, col).Value = hasil1
                    End If
                Next
            Next
            
            For Row = 3 To lv2.ListItems.Count + 2
                For col = 1 To lv2.ColumnHeaders.Count
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = "'" + lv2.ListItems(1).text
                    Else
                        '" 'cararandy 29032016 "
                        'Dim hasil1 As String
                            hasil1 = lv2.ListItems(1).SubItems(col - 1)
                            objExcelSheet.Cells(Row, col).Value = hasil1
                    End If
                Next
            Next
            
            For Row = 4 To lv3.ListItems.Count + 3
                For col = 1 To lv3.ColumnHeaders.Count
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = "'" + lv3.ListItems(1).text
                    Else
                        '" 'cararandy 29032016 "
                        'Dim hasil1 As String
                            hasil1 = lv3.ListItems(1).SubItems(col - 1)
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
'zzz:
            MsgBox "No data to export", vbInformation, Me.Caption
        End If

End Sub

Private Sub Command5_Click()
    Call lvv(lv4, TDBDate7.Value, TDBDate7.Value)
End Sub

Private Sub Command6_Click()
    Call lvv(lv5, TDBDate9.Value, TDBDate10.Value)
End Sub

Private Sub Command7_Click()
        Dim objExcel As New Excel.Application
        Dim objExcelSheet As Excel.Worksheet
        Dim col, Row As Integer
        Dim a As String
        'On Error GoTo zzz
        If lv4.ListItems.Count > 0 Then
            objExcel.Workbooks.ADD
            Set objExcelSheet = objExcel.Worksheets.ADD
         
        
            For col = 1 To lv4.ColumnHeaders.Count
                objExcelSheet.Cells(1, col).Value = lv4.ColumnHeaders(col)
            Next
         
            For Row = 2 To lv4.ListItems.Count + 1
                For col = 1 To lv4.ColumnHeaders.Count
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = "'" + lv4.ListItems(Row - 1).text
                    Else
                        '" 'cararandy 29032016 "
                        Dim hasil1 As String
                            hasil1 = lv4.ListItems(Row - 1).SubItems(col - 1)
                            objExcelSheet.Cells(Row, col).Value = hasil1
                    End If
                Next
            Next
            
            a = lv4.ListItems.Count


            B = 1
            For Row = Row To Row + lv5.ListItems.Count - 1
                For col = 1 To lv5.ColumnHeaders.Count
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = "'" + lv5.ListItems(col).text
                    Else
                        '" 'cararandy 29032016 "
                        'Dim hasil1 As String
                        If B <= lv5.ListItems.Count Then
                            hasil1 = lv5.ListItems(B).SubItems(col - 1)
                            objExcelSheet.Cells(Row, col).Value = hasil1
                        End If
                    End If
                Next
            B = B + 1
            Next
                     
            objExcelSheet.Columns.AutoFit
            CD_save.ShowOpen
            a = CD_save.FileName
         
            objExcelSheet.SaveAs a & ".xls"
            MsgBox "Export Completed", vbInformation, Me.Caption
         
            objExcel.Workbooks.Open a & ".xls"
            objExcel.Visible = True
        Else
'zzz:
            MsgBox "No data to export", vbInformation, Me.Caption
        End If


End Sub

Private Sub frmreportproductivity_DblClick()

End Sub

Private Sub Command8_Click()
        Dim objExcel As New Excel.Application
        Dim objExcelSheet As Excel.Worksheet
        Dim col, Row As Integer
        Dim a As String
        'On Error GoTo zzz
        If lv6.ListItems.Count > 0 Then
            objExcel.Workbooks.ADD
            Set objExcelSheet = objExcel.Worksheets.ADD
         
        
            For col = 1 To lv6.ColumnHeaders.Count
                objExcelSheet.Cells(1, col).Value = lv6.ColumnHeaders(col)
            Next
         
            For Row = 2 To lv6.ListItems.Count + 1
                For col = 1 To lv6.ColumnHeaders.Count
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = "'" + lv6.ListItems(Row - 1).text
                    Else
                        '" 'cararandy 29032016 "
                        Dim hasil1 As String
                            hasil1 = lv6.ListItems(Row - 1).SubItems(col - 1)
                            objExcelSheet.Cells(Row, col).Value = hasil1
                    End If
                Next
            Next
            
'            a = lv4.ListItems.Count
'
'
'            B = 1
'            For Row = Row To Row + lv5.ListItems.Count - 1
'                For col = 1 To lv5.ColumnHeaders.Count
'                    If col = 1 Then
'                            objExcelSheet.Cells(Row, col).Value = "'" + lv5.ListItems(col).text
'                    Else
'                        '" 'cararandy 29032016 "
'                        'Dim hasil1 As String
'                        If B <= lv5.ListItems.Count Then
'                            hasil1 = lv5.ListItems(B).SubItems(col - 1)
'                            objExcelSheet.Cells(Row, col).Value = hasil1
'                        End If
'                    End If
'                Next
'            B = B + 1
'            Next
                     
            objExcelSheet.Columns.AutoFit
            CD_save.ShowOpen
            a = CD_save.FileName
         
            objExcelSheet.SaveAs a & ".xls"
            MsgBox "Export Completed", vbInformation, Me.Caption
         
            objExcel.Workbooks.Open a & ".xls"
            objExcel.Visible = True
        Else
'zzz:
            MsgBox "No data to export", vbInformation, Me.Caption
        End If

End Sub

Private Sub Command9_Click()
    Call lvvv(lv6, TDBDate11.Value, TDBDate11.Value)
End Sub

Private Sub Form_Load()
    'SSTab1.TabVisible(1) = False
End Sub

Private Sub ListView4_DblClick()
    If ListView4.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = ListView4.SelectedItem.ListSubItems(3)
        Me.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If

End Sub

Private Sub lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lv1.SortKey = ColumnHeader.index - 1
    lv1.Sorted = True
End Sub

Private Sub lv4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lv4.SortKey = ColumnHeader.index - 1
    lv4.Sorted = True
End Sub

Private Sub lv6_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lv6.SortKey = ColumnHeader.index - 1
    lv6.Sorted = True
End Sub
