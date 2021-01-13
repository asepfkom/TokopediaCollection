VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form form_trade 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trading Form"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   20400
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   20400
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log Trade"
      Height          =   3555
      Left            =   120
      TabIndex        =   27
      Top             =   7440
      Width           =   20295
      Begin VB.CommandButton Command8 
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0000FF00&
         Caption         =   "Search"
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
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "form_trade.frx":0000
         Left            =   12720
         List            =   "form_trade.frx":0002
         TabIndex        =   31
         Text            =   "Combo1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "form_trade.frx":0004
         Left            =   12720
         List            =   "form_trade.frx":0006
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3150
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   5556
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
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   375
         Left            =   12720
         TabIndex        =   33
         Top             =   1200
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":0008
         Caption         =   "form_trade.frx":0120
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":018C
         Keys            =   "form_trade.frx":01AA
         Spin            =   "form_trade.frx":0208
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
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
         Format          =   "dd, mmm yyyy"
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
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate2 
         Height          =   375
         Left            =   15120
         TabIndex        =   34
         Top             =   1200
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":0230
         Caption         =   "form_trade.frx":0348
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":03B4
         Keys            =   "form_trade.frx":03D2
         Spin            =   "form_trade.frx":0430
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
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
         Format          =   "dd, mmm yyyy"
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
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin VB.Label Label9 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14640
         TabIndex        =   36
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "DateTrade :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   35
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Status   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   32
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   15
      Left            =   6000
      TabIndex        =   26
      Top             =   9720
      Width           =   135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data on TL"
      Height          =   7215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Autosms"
         Height          =   735
         Left            =   7440
         TabIndex        =   87
         Top             =   1440
         Width           =   1455
         Begin VB.CommandButton Command11 
            Caption         =   "Auto Sms"
            Height          =   375
            Left            =   240
            TabIndex        =   88
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   3720
         TabIndex        =   67
         Text            =   "Combo7"
         Top             =   1320
         Width           =   1455
      End
      Begin TDBDate6Ctl.TDBDate TDBDate3 
         Height          =   375
         Left            =   3720
         TabIndex        =   56
         Top             =   360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":0458
         Caption         =   "form_trade.frx":0570
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":05DC
         Keys            =   "form_trade.frx":05FA
         Spin            =   "form_trade.frx":0658
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   2.18403159439546E-315
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   375
         Left            =   7080
         TabIndex        =   51
         Top             =   360
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":0680
         Caption         =   "form_trade.frx":06A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":070C
         Keys            =   "form_trade.frx":072A
         Spin            =   "form_trade.frx":0774
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check3"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   255
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000E&
         Caption         =   "AGENT"
         ClipControls    =   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   45
         Top             =   960
         Width           =   2175
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "form_trade.frx":079C
            Left            =   1080
            List            =   "form_trade.frx":07A3
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label10 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Agent  :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000E&
         Caption         =   "TEAM"
         ClipControls    =   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   42
         Top             =   240
         Width           =   2175
         Begin VB.ComboBox cmb_team 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "form_trade.frx":07AC
            Left            =   1080
            List            =   "form_trade.frx":07AE
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl_team 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Team  :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "form_trade.frx":07B0
         Left            =   3960
         List            =   "form_trade.frx":07B2
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H8000000E&
         Caption         =   "Check2"
         Height          =   255
         Left            =   3720
         TabIndex        =   38
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Text            =   "0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek"
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
         Left            =   1320
         TabIndex        =   23
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "View"
         Height          =   375
         Left            =   9120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   10
         Top             =   1845
         Width           =   855
      End
      Begin VB.CommandButton cmd_trade 
         BackColor       =   &H0000FF00&
         Caption         =   "Trade"
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
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CheckBox cek_all_ptp 
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
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4350
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7673
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
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   375
         Left            =   8760
         TabIndex        =   52
         Top             =   360
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":07B4
         Caption         =   "form_trade.frx":07D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":0840
         Keys            =   "form_trade.frx":085E
         Spin            =   "form_trade.frx":08A8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate TDBDate4 
         Height          =   375
         Left            =   5160
         TabIndex        =   57
         Top             =   360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":08D0
         Caption         =   "form_trade.frx":09E8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":0A54
         Keys            =   "form_trade.frx":0A72
         Spin            =   "form_trade.frx":0AD0
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   2.18403159439546E-315
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate5 
         Height          =   375
         Left            =   3720
         TabIndex        =   60
         Top             =   840
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":0AF8
         Caption         =   "form_trade.frx":0C10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":0C7C
         Keys            =   "form_trade.frx":0C9A
         Spin            =   "form_trade.frx":0CF8
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   2.18403159439546E-315
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate6 
         Height          =   375
         Left            =   5160
         TabIndex        =   61
         Top             =   840
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":0D20
         Caption         =   "form_trade.frx":0E38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":0EA4
         Keys            =   "form_trade.frx":0EC2
         Spin            =   "form_trade.frx":0F20
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   2.18403159439546E-315
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber3 
         Height          =   375
         Left            =   7080
         TabIndex        =   63
         Top             =   840
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":0F48
         Caption         =   "form_trade.frx":0F68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":0FD4
         Keys            =   "form_trade.frx":0FF2
         Spin            =   "form_trade.frx":103C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber4 
         Height          =   375
         Left            =   8760
         TabIndex        =   64
         Top             =   840
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":1064
         Caption         =   "form_trade.frx":1084
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":10F0
         Keys            =   "form_trade.frx":110E
         Spin            =   "form_trade.frx":1158
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2880
         TabIndex        =   68
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Batch:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   66
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   65
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "LPA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6240
         TabIndex        =   62
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   59
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "LP Date:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   58
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   55
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "WO Date:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   54
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8400
         TabIndex        =   53
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Curbal:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   50
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbldata 
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
         Left            =   120
         TabIndex        =   6
         Top             =   6720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data on Trade"
      Height          =   7215
      Left            =   10320
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   7320
         TabIndex        =   86
         Text            =   "Combo7"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080FF80&
         Caption         =   "Go To TRADEFW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8880
         MaskColor       =   &H0080C0FF&
         TabIndex        =   41
         Top             =   120
         Width           =   1095
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
         TabIndex        =   25
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000FF00&
         Caption         =   "Refresh"
         Height          =   375
         Left            =   8880
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   22
         Top             =   840
         Width           =   1095
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
         TabIndex        =   21
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "Clear"
         Height          =   375
         Left            =   8880
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "form_trade.frx":1180
         Left            =   1080
         List            =   "form_trade.frx":1182
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Filter"
         Height          =   375
         Left            =   8880
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "form_trade.frx":1184
         Left            =   1080
         List            =   "form_trade.frx":1186
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   360
         Width           =   1215
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
         TabIndex        =   8
         Top             =   6600
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4350
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   7673
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
         Left            =   8280
         TabIndex        =   1
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate tgl_mulai1 
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   480
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":1188
         Caption         =   "form_trade.frx":12A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":130C
         Keys            =   "form_trade.frx":132A
         Spin            =   "form_trade.frx":1388
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
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
         Format          =   "dd, mmm yyyy"
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
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tgl_akhir1 
         Height          =   375
         Left            =   6000
         TabIndex        =   17
         Top             =   480
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":13B0
         Caption         =   "form_trade.frx":14C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":1534
         Keys            =   "form_trade.frx":1552
         Spin            =   "form_trade.frx":15B0
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
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
         Format          =   "dd, mmm yyyy"
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
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber5 
         Height          =   375
         Left            =   4200
         TabIndex        =   70
         Top             =   1200
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":15D8
         Caption         =   "form_trade.frx":15F8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":1664
         Keys            =   "form_trade.frx":1682
         Spin            =   "form_trade.frx":16CC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber6 
         Height          =   375
         Left            =   5760
         TabIndex        =   71
         Top             =   1200
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":16F4
         Caption         =   "form_trade.frx":1714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":1780
         Keys            =   "form_trade.frx":179E
         Spin            =   "form_trade.frx":17E8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate TDBDate7 
         Height          =   375
         Left            =   1080
         TabIndex        =   75
         Top             =   1200
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":1810
         Caption         =   "form_trade.frx":1928
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":1994
         Keys            =   "form_trade.frx":19B2
         Spin            =   "form_trade.frx":1A10
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   1.83561907008957E-315
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate8 
         Height          =   375
         Left            =   2520
         TabIndex        =   76
         Top             =   1200
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":1A38
         Caption         =   "form_trade.frx":1B50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":1BBC
         Keys            =   "form_trade.frx":1BDA
         Spin            =   "form_trade.frx":1C38
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   1.83561907008957E-315
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate9 
         Height          =   375
         Left            =   1080
         TabIndex        =   79
         Top             =   1680
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":1C60
         Caption         =   "form_trade.frx":1D78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":1DE4
         Keys            =   "form_trade.frx":1E02
         Spin            =   "form_trade.frx":1E60
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   1.83561907008957E-315
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate10 
         Height          =   375
         Left            =   2520
         TabIndex        =   80
         Top             =   1680
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":1E88
         Caption         =   "form_trade.frx":1FA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":200C
         Keys            =   "form_trade.frx":202A
         Spin            =   "form_trade.frx":2088
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   1.83561907008957E-315
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber7 
         Height          =   375
         Left            =   4200
         TabIndex        =   81
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":20B0
         Caption         =   "form_trade.frx":20D0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":213C
         Keys            =   "form_trade.frx":215A
         Spin            =   "form_trade.frx":21A4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber8 
         Height          =   375
         Left            =   5760
         TabIndex        =   82
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         Calculator      =   "form_trade.frx":21CC
         Caption         =   "form_trade.frx":21EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":2258
         Keys            =   "form_trade.frx":2276
         Spin            =   "form_trade.frx":22C0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Batch:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   8160
         TabIndex        =   85
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Curbal:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   3600
         TabIndex        =   84
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   5400
         TabIndex        =   83
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   2160
         TabIndex        =   78
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "LP Date  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   77
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   74
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "WO Date:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   73
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5400
         TabIndex        =   72
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Curbal:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3600
         TabIndex        =   69
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "TRADE DATE   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Status  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Team   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
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
         Left            =   8040
         TabIndex        =   9
         Top             =   6720
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "form_trade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cek_all_ptp_Click()
    Dim r As Integer
        
    If cek_all_ptp.Value = vbChecked Then
        If ListView1.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = True
        Next r
    Else
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = False
        Next r
    End If
End Sub

Private Sub isitrade()
    Dim query, query1, query3 As String
    Dim rs As ADODB.Recordset
    
    ListView1.ListItems.clear
    lbldata.Caption = ""
    
    ListView2.ListItems.clear
    
    query = "select mgm.custid,mgm.name,mgm.agentlama,mgm.f_cek_new,temp_trade.tanggal_trader,temp_trade.batch,temp_trade.wo_date,temp_trade.curbal,temp_trade.lpd,temp_trade.lpa from mgm,temp_trade where mgm.custid = temp_trade.custid and mgm.agent = 'TRADE' order by lpd"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    query1 = "delete from temp_trade where custid not in (select mgm.custid from mgm,temp_trade where mgm.custid = temp_trade.custid and mgm.agent = 'TRADE' order by custid)"
    M_OBJCONN.Execute query1
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = IIf(IsNull(rs("agentlama")), "", rs("agentlama"))
             listItem.SubItems(3) = cnull(rs("f_cek_new"))
             listItem.SubItems(4) = Format(cnull(rs("tanggal_trader")), "YYYY-MM-DD")
             listItem.SubItems(5) = cnull(rs("batch"))
             listItem.SubItems(6) = Format(IIf(IsNull(rs("wo_date")), "", rs("wo_date")), "yyyy-mm-dd")
             listItem.SubItems(7) = cnull(rs("curbal"))
             listItem.SubItems(8) = Format(cnull(rs("lpd")), "yyyy-mm-dd")
             listItem.SubItems(9) = cnull(rs("lpa"))
        rs.MoveNext
    Wend
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount

    Set rs = Nothing
    
End Sub

Private Sub isicombolv3()
    Combo3.clear
    Combo4.clear

    query = " select distinct(agent) from tbl_hst_trade"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Combo3.AddItem rs!agent
        rs.MoveNext
    Wend
    
    Set rs = Nothing
    
    query = " select distinct(status_account) from tbl_hst_trade where status_account <> ''"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    While Not rs.EOF
        Combo4.AddItem rs!status_account
        rs.MoveNext
    Wend
    
    Set rs = Nothing
    
End Sub

Private Sub searchinglv3()
    ListView3.ListItems.clear
    
    a = Format(TDBDate1.Value, "yyyy-mm-dd")
    B = Format(TDBDate2.Value, "yyyy-mm-dd")
    
    query3 = "select * from tbl_hst_trade where 1=1 "
    If Combo3.text <> "" Then
        query3 = query3 + " and agent = '" + Combo3.text + "' "
    End If
    If Combo4.text <> "" Then
        query3 = query3 + " and status_account = '" + Combo4.text + "' "
    End If
    If cnull(a) <> "" And cnull(B) <> "" Then
        query3 = query3 + " and date(tanggal_auto_trade) between '" + a + "' and '" + B + "'"
    End If
    query3 = query3 + " order by tanggal_auto_trade desc"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView3.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = IIf(IsNull(rs("agent")), "", rs("agent"))
             listItem.SubItems(3) = cnull(rs("status_account"))
             listItem.SubItems(4) = Format(cnull(rs("tanggal_trader")), "YYYY-MM-DD")
             listItem.SubItems(5) = Format(cnull(rs("tanggal_auto_trade")), "YYYY-MM-DD")
             listItem.SubItems(6) = cnull(rs("agent_kirim"))
        rs.MoveNext
    Wend
    
    Set rs = Nothing
End Sub

Private Sub filtertrade()
    Dim query As String
    Dim rs As ADODB.Recordset

    ListView2.ListItems.clear
    
    tgl1 = Format(tgl_mulai1.Value, "YYYY-MM-DD")
    tgl2 = Format(tgl_akhir1.Value, "YYYY-MM-DD")
    
    query = "select mgm.custid,mgm.name,mgm.agentlama,mgm.f_cek_new,temp_trade.tanggal_trader,temp_trade.batch,temp_trade.wo_date,temp_trade.curbal,temp_trade.lpd,temp_trade.lpa from mgm,temp_trade where mgm.custid = temp_trade.custid and mgm.agent = 'TRADE' "
    If Combo1.text <> "" Then
        query = query + " and mgm.agentlama = '" + Combo1.text + "' "
    End If
    If Combo2.text <> "" Then
        query = query + " and mgm.f_cek_new = '" + Combo2.text + "' "
    End If
    If tgl1 <> "" And tgl2 <> "" Then
        query = query + " and date(temp_trade.tanggal_trader) between '" + tgl1 + "' and '" + tgl2 + "' "
    End If
    If Combo8.text <> "" Then
        query = query + " and batch = '" + Combo8.text + "' "
    End If
    If TDBDate7.text <> "__/__/____" And TDBDate8.text <> "__/__/____" Then
        a = Format(TDBDate7.Value, "yyyy-mm-dd")
        a1 = Format(TDBDate8.Value, "yyyy-mm-dd")
        query = query + " and date(wo_date) between '" + a + "' and '" + a1 + "' "
    End If
    If TDBDate9.text <> "__/__/____" And TDBDate10.text <> "__/__/____" Then
        a = Format(TDBDate9.Value, "yyyy-mm-dd")
        a1 = Format(TDBDate10.Value, "yyyy-mm-dd")
        query = query + " and date(lpd) between '" + a + "' and '" + a1 + "' "
    End If
    If TDBNumber5.Value <> 0 Or TDBNumber6.Value <> 0 Then
        a = TDBNumber5.Value
        a1 = TDBNumber6.Value
        query = query + " and temp_trade.curbal >= " & a & " and temp_trade.curbal <= " & a1 & " "
    End If
    If TDBNumber7.Value <> 0 Or TDBNumber8.Value <> 0 Then
        query = query + " and trim(to_char(LastPay::numeric,'9999999999'))::bigint >= " & TDBNumber7.Value & " and trim(to_char(LastPay::numeric,'9999999999'))::bigint <= " & TDBNumber8.Value & " "
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query + " order by lpd", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , rs("custid"))
             listItem.SubItems(1) = rs("name")
             listItem.SubItems(2) = IIf(IsNull(rs("agentlama")), "", rs("agentlama"))
             listItem.SubItems(3) = cnull(rs("f_cek_new"))
             listItem.SubItems(4) = Format(rs("tanggal_trader"), "YYYY-MM-DD")
             listItem.SubItems(5) = cnull(rs("batch"))
             listItem.SubItems(6) = Format(IIf(IsNull(rs("wo_date")), "", rs("wo_date")), "yyyy-mm-dd")
             listItem.SubItems(7) = cnull(rs("curbal"))
             listItem.SubItems(8) = Format(cnull(rs("lpd")), "yyyy-mm-dd")
             listItem.SubItems(9) = cnull(rs("lpa"))
        rs.MoveNext
    Wend
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount
End Sub

Private Sub combo1dah()
    Dim query As String
    Dim rs As ADODB.Recordset

    Combo1.clear
    
    query = "select distinct (agentlama) from mgm,temp_trade where mgm.custid = temp_trade.custid order by 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Combo1.AddItem rs!agentlama
        rs.MoveNext
    Wend
    
    Combo2.clear
    
    query = "select distinct (f_cek_new) from mgm,temp_trade where mgm.custid = temp_trade.custid and f_cek_new <> '' order by 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Combo2.AddItem cnull(rs!f_cek_new)
        rs.MoveNext
    Wend
End Sub

Private Sub Check1_Click()
    Dim abc As String
    If Check1.Value = 1 Then
        If IsNumeric(Text1.text) Then
            abc = Text1.text
        Else
            MsgBox "Numeric Only"
            Text1.text = ""
        End If
        If ListView1.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
            
        For r = 1 To abc
            ListView1.ListItems(r).Checked = True
        Next r
    Else
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = False
        Next r
    End If
        
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
    
        If Combo5.text = "" Then
            MsgBox "Pilih Status!"
            Exit Sub
        End If
        
        Dim query As String
        Dim rs As ADODB.Recordset
        
        ListView1.ListItems.clear
        
        query = "select * from mgm  "
        If Check3.Value = 1 Then
            If cmb_team.text = "ALL" Then
                query = query + " where agent ilike 'TL%'"
            Else
                query = query + " where agent = '" + cmb_team.text + "'"
            End If
        ElseIf Check4.Value = 1 Then
            query = query + " where agent ilike 'D%'"
        End If
        If Combo5.text = "<blank>" Then
            query = query + "  and (f_cek_new = '' or f_cek_new is null)  order by custid "
        Else
            query = query + "  and f_cek_new = '" + Combo5.text + "' order by custid "
        End If
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not rs.EOF
            Set listItem = ListView1.ListItems.ADD(, , cnull(rs("custid")))
                 listItem.SubItems(1) = cnull(rs("name"))
                 listItem.SubItems(2) = cnull(rs("f_cek_new"))
                 listItem.SubItems(3) = IIf(IsNull(rs("agent")), "", rs("agent"))
                 listItem.SubItems(4) = cnull(rs("recsource"))
                 listItem.SubItems(5) = Format(rs("b_d"), "yyyy-mm-dd")
                 listItem.SubItems(6) = Format(rs("curbal"), "#,#")
                 listItem.SubItems(7) = Format(rs("Pay_Dt"), "yyyy-mm-dd")
                 listItem.SubItems(8) = Format(rs("LastPay"), "#,#")
            rs.MoveNext
        Wend
        
        lbldata.Caption = "Jumlah Data  : " & rs.RecordCount
        
        Set rs = Nothing
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Check4.Value = 0
        Frame5.Enabled = True
        Frame6.Enabled = False
    Else
        Frame5.Enabled = False
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Check3.Value = 0
        Frame6.Enabled = True
        Frame5.Enabled = False
    Else
        Frame6.Enabled = False
    End If
End Sub

Private Sub pretelan()
    q = "select distinct recsource from mgm order by 1"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseClient
    r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Combo7.clear
        
    While Not r.EOF
        Combo7.AddItem r!RECSOURCE
        r.MoveNext
    Wend
    
    qupd = "update temp_trade set batch = mgm.recsource, wo_date = mgm.b_d, curbal = mgm.curbal, lpd = mgm.Pay_Dt, lpa = mgm.LastPay from mgm where temp_trade.custid = mgm.custid and temp_trade.batch != mgm.recsource"
    M_OBJCONN.Execute qupd
    
    q1 = "select distinct batch from temp_trade order by 1"
    Set r1 = New ADODB.Recordset
    r1.CursorLocation = adUseClient
    r1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Combo8.clear
        
    While Not r1.EOF
        Combo8.AddItem r1!batch
        r1.MoveNext
    Wend
        
End Sub


Private Sub cmd_trade_Click()
    Dim K, w, cek As Integer
    Dim query As String
    Dim rs As ADODB.Recordset
    If cmb_team.text = "" And Combo6.text = "" Then
        MsgBox "Pilih team!"
        Exit Sub
    End If
    
    For K = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If

    For w = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(w).Checked = True Then
            CustId = ListView1.ListItems(w).text
            Nama = ListView1.ListItems(w).ListSubItems(1)
            If Check4.Value = 1 Then
                dc = ListView1.ListItems(w).ListSubItems(3)
                q = "select team from usertbl where userid = '" + dc + "'"
                Set rs = New ADODB.Recordset
                rs.CursorLocation = adUseClient
                rs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                agent = rs!TEAM
            ElseIf Check3.Value = 1 Then
                agent = ListView1.ListItems(w).ListSubItems(3)
            End If
            STATUS = ListView1.ListItems(w).ListSubItems(2)
            RECSOURCE = ListView1.ListItems(w).ListSubItems(4)
            wodate = Format(ListView1.ListItems(w).ListSubItems(5), "yyyy-mm-dd")
            cur_bal = Format(ListView1.ListItems(w).ListSubItems(6), "##")
            If cur_bal = "" Then
                curbal = 0
            End If
            'cur_bal = IIf(IsNull(cur_bal1), 0, cur_bal1)
            LP_D = Format(ListView1.ListItems(w).ListSubItems(7), "yyyy-mm-dd")
            If LP_D = "" Then
                LP_D = "1700-09-09"
            End If
            lp_a = Format(ListView1.ListItems(w).ListSubItems(8), "##")
            If lp_a = "" Then
                lp_a = 0
            End If
            'LP_A = IIf(IsNull(LP_A1), 0, LP_A1)
            
            query = "INSERT INTO temp_trade (custid,name,agent,status_account,batch,wo_date,curbal,lpd,lpa) values ('" + CustId + "','" + Nama + "','" + agent + "','" + STATUS + "', '" + RECSOURCE + "', '" + wodate + "'," & cur_bal & ",'" + LP_D + "'," & lp_a & ") ;" & vbCrLf
            If Check4.Value = 1 Then
                query = query + "UPDATE mgm set agent = 'TRADE', agentlama = '" + agent + "' where custid = '" + CustId + "' ;"
            ElseIf Check3.Value = 1 Then
                query = query + "UPDATE mgm set agent = 'TRADE', agentlama = agent where custid = '" + CustId + "' ;"
            End If
            M_OBJCONN.Execute query
        End If
    Next w
    
    
    MsgBox "Data Traded"
    
    
    
    ListView2.ListItems.clear
    
    query = "select mgm.custid,mgm.name,mgm.agentlama,mgm.f_cek_new,temp_trade.tanggal_trader,temp_trade.batch,temp_trade.wo_date,temp_trade.curbal,temp_trade.lpd,temp_trade.lpa from mgm,temp_trade where mgm.custid = temp_trade.custid and mgm.agent = 'TRADE' order by lpd"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = IIf(IsNull(rs("agentlama")), "", rs("agentlama"))
             listItem.SubItems(3) = cnull(rs("f_cek_new"))
             listItem.SubItems(4) = Format(cnull(rs("tanggal_trader")), "yyyy-mm-dd")
             listItem.SubItems(5) = cnull(rs("batch"))
             listItem.SubItems(6) = Format(IIf(IsNull(rs("wo_date")), "", rs("wo_date")), "yyyy-mm-dd")
             listItem.SubItems(7) = cnull(rs("curbal"))
             listItem.SubItems(8) = Format(cnull(rs("lpd")), "yyyy-mm-dd")
             listItem.SubItems(9) = cnull(rs("lpa"))
        rs.MoveNext
    Wend
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount
    
    
    cek_all_ptp.Value = 0
    Call showmgmtl
    Call showtl
    Call Form_Load
End Sub

Private Sub Command1_Click()
    If cmb_team.text = "" And Combo6.text = "" Then
        MsgBox "Pilih Data!"
        Exit Sub
    End If
    Check2.Value = 0
    Call showmgmtl
End Sub

Private Sub Command10_Click()
    formtradefwo.Show 1
End Sub

Private Sub Command11_Click()
    cek = 0
    For K = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For w = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(w).Checked = True Then
            CustId = ListView1.ListItems(w).text
        
            qupdate = "update mgm set autosms = 1 where custid = '" & CustId & "'"
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
    Call filtertrade
End Sub

Private Sub Command4_Click()
    Combo1.text = ""
    Combo2.text = ""
    tgl_mulai1.Value = Null
    tgl_akhir1.Value = Null
    TDBDate7.Value = Null
    TDBDate8.Value = Null
    TDBDate9.Value = Null
    TDBDate10.Value = Null
    TDBNumber5.Value = Null
    TDBNumber6.Value = Null
    TDBNumber7.Value = Null
    TDBNumber8.Value = Null
    Combo7.text = ""
End Sub

Private Sub autotrade()
    Dim query, query1, cmdsql, CMDSQL1, abc As String
    Dim a, B, c, d, e As Integer
    Dim rs, rs1, RS2, rs3, rs4, rs5 As ADODB.Recordset
    
    
    query = "select * from temp_trade where 1 = 1"
    If Combo1.text <> "" Then
        query = query + " and agent = '" + Combo1.text + "' "
    End If
    If Combo2.text <> "" Then
        query = query + " and status_account = '" + Combo2.text + "' "
    End If
    If tgl1 <> "" And tgl2 <> "" Then
        query = query + " and date(temp_trade.tanggal_trader) between '" + tgl1 + "' and '" + tgl2 + "' "
    End If
        If Combo8.text <> "" Then
        query = query + " and batch = '" + Combo8.text + "' "
    End If
    If TDBDate7.text <> "__/__/____" And TDBDate8.text <> "__/__/____" Then
        a = Format(TDBDate7.Value, "yyyy-mm-dd")
        a1 = Format(TDBDate8.Value, "yyyy-mm-dd")
        query = query + " and date(wo_date) between '" + a + "' and '" + a1 + "' "
    End If
    If TDBDate9.text <> "__/__/____" And TDBDate10.text <> "__/__/____" Then
        a = Format(TDBDate9.Value, "yyyy-mm-dd")
        a1 = Format(TDBDate10.Value, "yyyy-mm-dd")
        query = query + " and date(lpd) between '" + a + "' and '" + a1 + "' "
    End If
    If TDBNumber5.Value <> 0 Or TDBNumber6.Value <> 0 Then
        a = TDBNumber5.Value
        a1 = TDBNumber6.Value
        query = query + " and curbal >= " & a & " and curbal <= " & a1 & " "
    End If
    If TDBNumber7.Value <> 0 Or TDBNumber8.Value <> 0 Then
        query = query + " and trim(to_char(LastPay::numeric,'9999999999'))::bigint >= " & TDBNumber7.Value & " and trim(to_char(LastPay::numeric,'9999999999'))::bigint <= " & TDBNumber8.Value & " "
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    jml_trade = rs.RecordCount
    
    '<28sept2017
'    query1 = "select a.*, b.team from ("
'    query1 = query1 + " select count(agent),agent from mgm where agent ilike 'D%' and agent not ilike 'Dece%' and agent not ilike 'DOD%' group by 2) a,"
'    query1 = query1 + " (select userid, team from usertbl where userid ilike 'D%' and userid not ilike 'Dece%' and userid not ilike 'DOD%') b where a.agent = b.userid order by 1"
'----------------------------

    '28sept2017
    query1 = " select count(agent), userid, team from ( " & vbCrLf
    query1 = query1 + " select userid,team from usertbl where userid ilike 'D%' and userid not ilike 'Dece%' and userid not ilike 'DOD%' and usertype = 1 and aktif = 0 order by 1) a left join " & vbCrLf
    query1 = query1 + " (select * from mgm where custid not in ( select custid from (" & vbCrLf
    query1 = query1 + " select custid, count(custid) regular from (" & vbCrLf
    query1 = query1 + " select custid, to_char(paydate, 'yyyy-mm'), sum(payment)  from tbllunas where to_char(paydate, 'yyyy-mm') >= to_char(now() - interval '3 month', 'yyyy-mm')  group by 1,2 " & vbCrLf
    query1 = query1 + " ) a group by 1 order by 2 desc ) abc where regular > 1) " & vbCrLf
    query1 = query1 + " ) mgm " & vbCrLf
    query1 = query1 + " on a.userid = mgm.agent group by 2,3 order by 1 " & vbCrLf
    '----------------------------
    
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open query1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    cmdsql = "select jml from dataperagent"
    Set RS2 = New ADODB.Recordset
    RS2.CursorLocation = adUseClient
    RS2.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    B = RS2!jml
    e = 0

    For a = 1 To rs1.RecordCount
        If rs1!Count < B Then
        c = B - rs1!Count
            
            tgl1 = Format(tgl_mulai1.Value, "YYYY-MM-DD")
            tgl2 = Format(tgl_akhir1.Value, "YYYY-MM-DD")
            
            abc = " select a.custid from mgm a, temp_trade b where a.custid = b.custid and b.agent != '" + rs1!TEAM + "' and a.agent = 'TRADE' "
            
                If Combo1.text <> "" Then
                    abc = abc + " and a.agentlama = '" + Combo1.text + "' "
                End If
                If Combo2.text <> "" Then
                    abc = abc + " and a.f_cek_new = '" + Combo2.text + "' "
                End If
                If tgl1 <> "" And tgl2 <> "" Then
                    abc = abc + " and date(temp_trade.tanggal_trader) between '" + tgl1 + "' and '" + tgl2 + "' "
                End If
            Set rs5 = New ADODB.Recordset
            rs5.CursorLocation = adUseClient
            rs5.Open abc & "order by 1 limit " & c, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            e = e + rs5.RecordCount
            
'            If rs5.RecordCount = 0 Then
'                GoTo bawah
'            End If
            
            query3 = "update mgm set agent = '" + rs1!Userid + "' where custid in ( "
            query3 = query3 + abc & "order by 1 limit " & c
            query3 = query3 + " ) "
            M_OBJCONN.Execute query3
                        
            For d = 1 To rs5.RecordCount
                query5 = "INSERT INTO tbl_hst_trade (custid,name,agent,status_account,tanggal_trader) select custid,name,agent,status_account,tanggal_trader from temp_trade where custid = '" + rs5!CustId + "' "
                M_OBJCONN.Execute query5
                
                query6 = "UPDATE tbl_hst_trade set agent_kirim = '" + rs1!Userid + "' where custid = '" + rs5!CustId + "' "
                M_OBJCONN.Execute query6
                
                query4 = "delete from temp_trade where custid = '" + rs5!CustId + "' "
                M_OBJCONN.Execute query4
                rs5.MoveNext
            Next d
        End If
        rs1.MoveNext
    Next a
    
    CMDSQL1 = "select * from mgm where agent = 'TRADE'"
    If Combo1.text <> "" Then
        CMDSQL1 = CMDSQL1 + " and agentlama = '" + Combo1.text + "' "
    End If
    If Combo2.text <> "" Then
        CMDSQL1 = CMDSQL1 + " and f_cek_new = '" + Combo2.text + "' "
    End If
    Set rs4 = New ADODB.Recordset
    rs4.CursorLocation = adUseClient
    rs4.Open CMDSQL1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If jml_trade = rs4.RecordCount Then
        MsgBox "Tidak ada data yang Ter-Trade"
    Else
'bawah:
        MsgBox "Data Ter-Trader sebanyak " & e
        Call Form_Load
    End If
End Sub

Private Sub Command5_Click()
    Call autotrade
End Sub

Private Sub Command6_Click()
    Call Form_Load
End Sub

Private Sub Command7_Click()
    If form_trade.Height = 7830 Then
        ListView3.ListItems.clear
    
        query3 = "select * from tbl_hst_trade  order by tanggal_auto_trade desc limit 1000"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open query3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not rs.EOF
            Set listItem = ListView3.ListItems.ADD(, , cnull(rs("custid")))
                 listItem.SubItems(1) = cnull(rs("name"))
                 listItem.SubItems(2) = IIf(IsNull(rs("agent")), "", rs("agent"))
                 listItem.SubItems(3) = cnull(rs("status_account"))
                 listItem.SubItems(4) = Format(cnull(rs("tanggal_trader")), "YYYY-MM-DD")
                 listItem.SubItems(5) = Format(cnull(rs("tanggal_auto_trade")), "YYYY-MM-DD")
                 listItem.SubItems(6) = cnull(rs("agent_kirim"))
            rs.MoveNext
        Wend
        
        Set rs = Nothing

    
        form_trade.Height = 11520
        Command7.Caption = "Hide Log Trade"
    Else
        form_trade.Height = 7830
        Command7.Caption = "Show Log Trade"
    End If
End Sub

Private Sub Command8_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If ListView3.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView3.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView3.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView3.ListItems.Count + 1
            For col = 1 To ListView3.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = "'" + ListView3.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = ListView3.ListItems(Row - 1).SubItems(col - 1)
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

Private Sub Command9_Click()
    Call searchinglv3
End Sub

Private Sub Form_Load()
    Call showtl
    Call header
    Call isitrade
    Call combo1dah
    Call isicombolv3
    Call pretelan
End Sub

Private Sub showtl()
    Dim query As String
    Dim rs As ADODB.Recordset
    
    cmb_team.clear
    
    query = "select distinct agent from mgm where agent ilike 'TL%' or (agent = 'JOKO' or agent = 'REZA')  order by 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    'cmb_team.AddItem "ONTARIO"
    cmb_team.AddItem "ALL"
    While Not rs.EOF
        cmb_team.AddItem rs!agent
        rs.MoveNext
    Wend
End Sub

Private Sub showmgmtl()
    Dim query As String
    Dim rs As ADODB.Recordset
    
    ListView1.ListItems.clear
    Combo5.clear
    
    If Check3.Value = 0 And Check4.Value = 0 Then
        MsgBox "Pilih Category yang ingin disearch"
        Exit Sub
    End If
    
        qu = "delete from temp_max;" & vbCrLf
        qu = qu + "insert into temp_max select max from ( select custid, max(id) from tbllunas group by 1 ) a;" & vbCrLf
        qu = qu + " delete from tbllunas_temp;" & vbCrLf
        qu = qu + " insert into tbllunas_temp select * from tbllunas where id in (select * from temp_max) "
        M_OBJCONN.Execute qu
        
    If Check3.Value = 1 Then
        If cmb_team.text = "ALL" Then
            query = "select * from mgm a left join tbllunas_temp b on a.custid = b.custid where a.agent ilike 'TL%' "
        Else
            query = "select * from mgm a left join tbllunas_temp b on a.custid = b.custid where a.agent = '" + cmb_team.text + "' "
        End If
    ElseIf Check4.Value = 1 Then
        query = "select * from mgm a left join tbllunas_temp b on a.custid = b.custid where a.agent ilike 'D%' "
    End If
    If Combo7.text <> "" Then
        query = query + " and recsource = '" + Combo7.text + "' "
    End If
    If TDBDate3.text <> "__/__/____" And TDBDate4.text <> "__/__/____" Then
        a = Format(TDBDate3.Value, "yyyy-mm-dd")
        a1 = Format(TDBDate4.Value, "yyyy-mm-dd")
        query = query + " and date(b_d) between '" + a + "' and '" + a1 + "' "
    End If
    If TDBDate5.text <> "__/__/____" And TDBDate6.text <> "__/__/____" Then
        a = Format(TDBDate5.Value, "yyyy-mm-dd")
        a1 = Format(TDBDate6.Value, "yyyy-mm-dd")
        query = query + " and date(paydate) between '" + a + "' and '" + a1 + "' "
    End If
    If TDBNumber1.Value <> 0 Or TDBNumber2.Value <> 0 Then
        a = TDBNumber1.Value
        a1 = TDBNumber2.Value
        query = query + " and curbal >= " & a & " and curbal <= " & a1 & " "
    End If
    If TDBNumber3.Value <> 0 Or TDBNumber4.Value <> 0 Then
        query = query + " and trim(to_char(payment::numeric,'9999999999'))::bigint >= " & TDBNumber3.Value & " and trim(to_char(payment::numeric,'9999999999'))::bigint <= " & TDBNumber4.Value & " "
    End If
        query = query + " and a.custid in ( "
        query = query + " select custid from ("
        query = query + " select custid, count(custid) regular from ("
        query = query + " select custid, to_char(paydate, 'yyyy-mm'), sum(payment)  from tbllunas where to_char(paydate, 'yyyy-mm') >= to_char(now() - interval '3 month', 'yyyy-mm')  group by 1,2"
        query = query + " ) a group by 1"
        query = query + " order by 2 desc ) abc"
        query = query + " ) "
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query & " order by a.custid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = cnull(rs("f_cek_new"))
             listItem.SubItems(3) = IIf(IsNull(rs("agent")), "", rs("agent"))
             listItem.SubItems(4) = cnull(rs("recsource"))
             listItem.SubItems(5) = cnull(Format(rs("b_d"), "yyyy-mm-dd"))
             listItem.SubItems(6) = cnull(Format(rs("curbal"), "#,#"))
             listItem.SubItems(7) = cnull(Format(rs("paydate"), "yyyy-mm-dd"))
             listItem.SubItems(8) = cnull(Format(rs("payment"), "#,#"))
             
             'paydate n payment
        rs.MoveNext
    Wend
    
    lbldata.Caption = "Jumlah Data  : " & rs.RecordCount
    
    Set rs = Nothing
        
    If Check3.Value = 1 Then
        Combo5.clear
        If cmb_team.text = "ALL" Then
            query = "select case when f_cek_new = '' or f_cek_new is null then '<blank>' else f_cek_new end from ( select distinct(f_cek_new) from mgm where agent ilike 'TL%' order by 1 ) a "
        Else
            query = "select case when f_cek_new = '' or f_cek_new is null then '<blank>' else f_cek_new end from ( select distinct(f_cek_new) from mgm where agent = '" + cmb_team.text + "' order by 1 ) a"
        End If
    ElseIf Check4.Value = 1 Then
        query = "select case when f_cek_new = '' or f_cek_new is null then '<blank>' else f_cek_new end from ( select distinct(f_cek_new) from mgm where agent ilike 'D%' order by 1 ) a "
    End If
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not rs.EOF
            Combo5.AddItem rs!f_cek_new
            rs.MoveNext
        Wend
        
        Set rs = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.clear
    ListView2.ColumnHeaders.clear
    ListView3.ColumnHeaders.clear

    ListView1.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "Status Account", 20 * 120
    ListView1.ColumnHeaders.ADD 4, , "Agent", 8 * 120
    ListView1.ColumnHeaders.ADD 5, , "Batch", 10 * 120
    ListView1.ColumnHeaders.ADD 6, , "WO DATE", 20 * 120
    ListView1.ColumnHeaders.ADD 7, , "Curbal", 20 * 120
    ListView1.ColumnHeaders.ADD 8, , "LPD", 8 * 120
    ListView1.ColumnHeaders.ADD 9, , "LPA", 8 * 120
       
    ListView2.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
    ListView2.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView2.ColumnHeaders.ADD 3, , "Agent", 8 * 120
    ListView2.ColumnHeaders.ADD 4, , "Status Account", 20 * 120
    ListView2.ColumnHeaders.ADD 5, , "Tanggal Trade", 20 * 120
    ListView2.ColumnHeaders.ADD 6, , "Batch", 10 * 120
    ListView2.ColumnHeaders.ADD 7, , "WO DATE", 20 * 120
    ListView2.ColumnHeaders.ADD 8, , "Curbal", 20 * 120
    ListView2.ColumnHeaders.ADD 9, , "LPD", 8 * 120
    ListView2.ColumnHeaders.ADD 10, , "LPA", 8 * 120
    
    ListView3.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
    ListView3.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView3.ColumnHeaders.ADD 3, , "Agent", 20 * 50
    ListView3.ColumnHeaders.ADD 4, , "Status Account", 20 * 50
    ListView3.ColumnHeaders.ADD 5, , "Tanggal Trade", 20 * 120
    ListView3.ColumnHeaders.ADD 6, , "Tanggal Auto Trade", 20 * 120
    ListView3.ColumnHeaders.ADD 7, , "Trade To", 20 * 50

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = ListView1.SelectedItem.text
        form_trade.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
End Sub

Private Sub ListView2_DblClick()
    If ListView2.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = ListView2.SelectedItem.text
        form_trade.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True

End Sub
