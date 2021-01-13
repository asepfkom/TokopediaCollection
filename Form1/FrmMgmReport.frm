VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMgmReport 
   Appearance      =   0  'Flat
   BackColor       =   &H00B1FDD5&
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4755
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   9000
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Report Talk Time"
      Height          =   375
      Left            =   8040
      TabIndex        =   39
      Top             =   4200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Excel"
      Height          =   360
      Index           =   2
      Left            =   8020
      TabIndex        =   37
      Top             =   3480
      Width           =   1005
   End
   Begin VB.TextBox txtcustid 
      Height          =   285
      Left            =   5640
      TabIndex        =   36
      Top             =   3540
      Width           =   2325
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   7110
      Top             =   2460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Regular Payment"
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   3975
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Iregular to PaidOff"
      Height          =   255
      Left            =   1800
      TabIndex        =   33
      Top             =   3975
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Regular to PaidOff"
      Height          =   255
      Left            =   3810
      TabIndex        =   32
      Top             =   3975
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox CmbCek 
      Height          =   315
      Left            =   9165
      TabIndex        =   24
      Top             =   3105
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   360
      Index           =   0
      Left            =   9060
      TabIndex        =   12
      Top             =   3480
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   10230
      TabIndex        =   13
      Top             =   3480
      Width           =   1125
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   7
      Top             =   1590
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   6570
      TabIndex        =   6
      Top             =   1560
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B8E2D4&
      Caption         =   "Choose One..."
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   5550
      TabIndex        =   14
      Top             =   420
      Width           =   5805
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   3630
         TabIndex        =   2
         Top             =   195
         Width           =   2130
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1215
         TabIndex        =   1
         Top             =   180
         Width           =   2085
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   1
         ItemData        =   "FrmMgmReport.frx":0000
         Left            =   3630
         List            =   "FrmMgmReport.frx":0002
         TabIndex        =   5
         Top             =   540
         Width           =   2130
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   0
         ItemData        =   "FrmMgmReport.frx":0004
         Left            =   1230
         List            =   "FrmMgmReport.frx":0006
         TabIndex        =   4
         Top             =   540
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Agent        :"
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Supervisor :"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "to"
         Height          =   300
         Index           =   2
         Left            =   3375
         TabIndex        =   18
         Top             =   225
         Width           =   270
      End
      Begin VB.Label Label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "to"
         Height          =   300
         Index           =   6
         Left            =   3405
         TabIndex        =   17
         Top             =   570
         Width           =   270
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4125
      Left            =   -30
      TabIndex        =   16
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   7276
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
      Appearance      =   0
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   15
      TabIndex        =   15
      Top             =   4365
      Visible         =   0   'False
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   10
      Top             =   1905
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport.frx":0008
      Caption         =   "FrmMgmReport.frx":0120
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport.frx":018C
      Keys            =   "FrmMgmReport.frx":01AA
      Spin            =   "FrmMgmReport.frx":0208
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   0
      Left            =   6570
      TabIndex        =   8
      Top             =   1905
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport.frx":0230
      Caption         =   "FrmMgmReport.frx":0348
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport.frx":03B4
      Keys            =   "FrmMgmReport.frx":03D2
      Spin            =   "FrmMgmReport.frx":0430
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   0
      Left            =   7950
      TabIndex        =   9
      Top             =   1920
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReport.frx":0458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport.frx":04C4
      Spin            =   "FrmMgmReport.frx":0514
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn:ss"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn:ss"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.999988425925926
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
      Text            =   "__:__:__"
      ValidateMode    =   0
      ValueVT         =   2010382337
      Value           =   2.12482692446619E-314
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   1
      Left            =   10485
      TabIndex        =   11
      Top             =   1905
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReport.frx":053C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport.frx":05A8
      Spin            =   "FrmMgmReport.frx":05F8
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
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
      Value           =   0.870289351851852
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   2
      Left            =   6600
      TabIndex        =   26
      Top             =   2535
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport.frx":0620
      Caption         =   "FrmMgmReport.frx":0738
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport.frx":07A4
      Keys            =   "FrmMgmReport.frx":07C2
      Spin            =   "FrmMgmReport.frx":0820
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   3
      Left            =   9120
      TabIndex        =   27
      Top             =   2535
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport.frx":0848
      Caption         =   "FrmMgmReport.frx":0960
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport.frx":09CC
      Keys            =   "FrmMgmReport.frx":09EA
      Spin            =   "FrmMgmReport.frx":0A48
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   2
      Left            =   7920
      TabIndex        =   28
      Top             =   2535
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReport.frx":0A70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport.frx":0ADC
      Spin            =   "FrmMgmReport.frx":0B2C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn:ss"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn:ss"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.999988425925926
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
      Text            =   "__:__:__"
      ValidateMode    =   0
      ValueVT         =   2010382337
      Value           =   2.12482692446619E-314
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   3
      Left            =   10440
      TabIndex        =   29
      Top             =   2535
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReport.frx":0B54
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport.frx":0BC0
      Spin            =   "FrmMgmReport.frx":0C10
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn:ss"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn:ss"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.999988425925926
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
      Text            =   "__:__:__"
      ValidateMode    =   0
      ValueVT         =   2010382337
      Value           =   2.12482692446619E-314
   End
   Begin VB.TextBox TxtPath 
      Height          =   285
      Left            =   0
      TabIndex        =   38
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Custid"
      Height          =   255
      Left            =   5670
      TabIndex        =   35
      Top             =   3330
      Width           =   645
   End
   Begin VB.Label Label1 
      BackColor       =   &H009AD6C2&
      Caption         =   "to"
      Height          =   300
      Index           =   7
      Left            =   8880
      TabIndex        =   31
      Top             =   2535
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009AD6C2&
      Caption         =   "Comparator Date :"
      Height          =   420
      Index           =   3
      Left            =   5640
      TabIndex        =   30
      Top             =   2415
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009AD6C2&
      Caption         =   "Status Cek :"
      Height          =   315
      Left            =   7875
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackColor       =   &H009AD6C2&
      Caption         =   "to"
      Height          =   255
      Index           =   1
      Left            =   8850
      TabIndex        =   23
      Top             =   1605
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009AD6C2&
      Caption         =   "From Batch :"
      Height          =   300
      Index           =   0
      Left            =   5595
      TabIndex        =   22
      Top             =   1620
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009AD6C2&
      Caption         =   "Date :"
      Height          =   300
      Index           =   5
      Left            =   5595
      TabIndex        =   21
      Top             =   1950
      Width           =   930
   End
   Begin VB.Label Label1 
      BackColor       =   &H009AD6C2&
      Caption         =   "to"
      Height          =   300
      Index           =   4
      Left            =   8850
      TabIndex        =   20
      Top             =   1935
      Width           =   270
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B1FDD5&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5565
      TabIndex        =   19
      Top             =   15
      Width           =   5745
   End
End
Attribute VB_Name = "FrmMgmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim M_Objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim JUMLAHVOL As String
Dim batch As String
Dim cmdsql As String
Dim CMDSQL1 As String
Dim STATUS As String
Dim LAmount As String
Dim LAgent As String
Dim LAgent1 As String
Dim Last As String
Dim jml As String
Dim Lf_cek As String
Dim Lvol As String
Dim ExlObj As Excel.Application      ' Create excel object
Dim b_excel As Boolean

Private Sub VolumeUtilized()
Dim m_hst As New ADODB.Recordset
'Dim tglawal As String
'Dim tglakhir As String
Dim m_msgbox As Variant

On Error GoTo eddder
'tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
'tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
'Set m_hst = New ADODB.Recordset
m_hst.CursorLocation = adUseClient

If Option1(0).Value Then
cmdsql = "SELECT mgm.agent,  sum(mgm.amountwo)as jmlAmount from mgm "
cmdsql = cmdsql + " where custid in (SELECT distinct custid from mgm_hst "
cmdsql = cmdsql + " where  tgl BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "') "
cmdsql = cmdsql + " and  RECSOURCE Between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'  "
cmdsql = cmdsql + " and agent between '" + Combo2(0).text + "'  and '" + Combo2(1).text + "' group by agent"
Else
If Option1(1).Value Then
cmdsql = "SELECT mgm.agent,  sum(mgm.amountwo)as jmlAmount from mgm where custid in "
cmdsql = cmdsql + "(SELECT distinct custid from mgm_hst where  "
cmdsql = cmdsql + " tgl BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "') "
cmdsql = cmdsql + " and  RECSOURCE Between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'  "
cmdsql = cmdsql + " and agent in(select userid from usertbl where "
cmdsql = cmdsql + " spvcode  >='" + Combo3(0).text + "' and spvcode <='" + Combo3(1).text + "')group by agent"
End If
End If

m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 2
While Not m_hst.EOF
ProgressBar1.Value = m_hst.Bookmark
LAgent1 = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
LAmount = Trim(CStr(IIf(IsNull(m_hst!jmlamount), 0, m_hst!jmlamount)))
CMDSQL1 = " Update TrackingRptPerPrgBatch set VOLUTILIZED= '" + LAmount + "' where AOC='" + LAgent1 + "'"
M_RPTCONN.execute CMDSQL1
m_hst.MoveNext
Wend
Set m_hst = Nothing
cmdsql = Empty
CMDSQL1 = Empty
LAgent1 = Empty

Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next
End Sub

Private Sub Hitung_Pergerakan_status()
    Dim M_Objrs_Result As New ADODB.Recordset
    Dim iQuery As String
    Dim Descol As String
    Dim Result_SK As Integer
    Dim Result_VL As Integer
    Dim Result_ON As Integer
    Dim Result_SP As Integer
    Dim Result_POP As Integer
    Dim Result_PR As Integer
    Dim Result_PO As Integer
    Dim Result_PTP_PO As Integer
    
    M_OBJCONN.execute "DROP TABLE IF EXISTS tbltemp_pergerakan_status"

    iQuery = " CREATE TABLE tbltemp_pergerakan_status AS "
    iQuery = iQuery + " SELECT * FROM ( "
    iQuery = iQuery + " SELECT * FROM ("
    iQuery = iQuery + " SELECT agent,custid, status_call_sebelum, f_cek_new as status_call_sekarang, tglcall"
    iQuery = iQuery + " FROM mgm where tglcall between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND agent in("
    iQuery = iQuery + " select userid from usertbl where spvcode >='" + Combo3(0).text + "' and spvcode <='" + Combo3(1).text + "' order by spvcode)) a"
    iQuery = iQuery + " left join ("
    iQuery = iQuery + " select jenis as jenis_sebelum,level_status as level_status_sebelum"
    iQuery = iQuery + " from contacteddesc) b on a.status_call_sebelum = b.jenis_sebelum) c"
    iQuery = iQuery + " left join ("
    iQuery = iQuery + " select jenis as jenis_sekarang,level_status as level_status_sekarang"
    iQuery = iQuery + " from contacteddesc) d on c.status_call_sekarang = d.jenis_sekarang  WHERE level_status_sekarang - level_status_sebelum <> 0 order by agent"
    
    M_OBJCONN.execute iQuery
    
    iQuery = "    select a1.agent,jumlah_sk,jumlah_vl,jumlah_on,jumlah_sp,jumlah_pop,jumlah_pr,jumlah_po,jumlah_ptp_po from ("
    iQuery = iQuery + " SELECT * FROM (SELECT agent, count(status_call_sekarang) as JUMLAH_SK from tbltemp_pergerakan_status where status_call_sekarang = 'SK-' group by agent) as a "
    iQuery = iQuery + " ) as a1 left join ("
    iQuery = iQuery + " SELECT agent, count(status_call_sekarang) as JUMLAH_VL from tbltemp_pergerakan_status where status_call_sekarang = 'VL-' group by agent"
    iQuery = iQuery + " ) as a2 on (a1.agent=a2.agent) left join ("
    iQuery = iQuery + " SELECT agent, count(status_call_sekarang) as JUMLAH_ON from tbltemp_pergerakan_status where status_call_sekarang = 'ON-' group by agent) as a3 on(a1.agent=a3.agent) left join ("
    iQuery = iQuery + " SELECT agent, count(status_call_sekarang) as JUMLAH_SP from tbltemp_pergerakan_status where status_call_sekarang = 'SP-' group by agent) as a4 on(a1.agent=a4.agent) left join ("
    iQuery = iQuery + " SELECT agent, count(status_call_sekarang) as JUMLAH_POP from tbltemp_pergerakan_status where status_call_sekarang = 'POP-' group by agent) as a5 on(a1.agent=a5.agent) left join ("
    iQuery = iQuery + " SELECT agent, count(status_call_sekarang) as JUMLAH_PR from tbltemp_pergerakan_status where status_call_sekarang = 'PR-' group by agent) as a6 on(a1.agent=a6.agent) left join ("
    iQuery = iQuery + " SELECT agent, count(status_call_sekarang) as JUMLAH_PO from tbltemp_pergerakan_status where status_call_sekarang = 'PO-' group by agent) as a7 on(a1.agent=a7.agent) left join ("
    iQuery = iQuery + " SELECT agent, count(status_call_sekarang) as JUMLAH_PTP_PO from tbltemp_pergerakan_status where status_call_sekarang = 'PTP-PO' group by agent) as a8 on(a1.agent=a8.agent)"
    
    Set M_Objrs_Result = New ADODB.Recordset
    M_Objrs_Result.CursorLocation = adUseClient
    M_Objrs_Result.Open iQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    

    While Not M_Objrs_Result.EOF
        Descol = Trim(IIf(IsNull(M_Objrs_Result!agent), "", M_Objrs_Result!agent))
        Result_SK = Trim(IIf(IsNull(M_Objrs_Result!jumlah_sk), 0, M_Objrs_Result!jumlah_sk))
        Result_VL = Trim(IIf(IsNull(M_Objrs_Result!jumlah_vl), 0, M_Objrs_Result!jumlah_vl))
        Result_ON = Trim(IIf(IsNull(M_Objrs_Result!jumlah_on), 0, M_Objrs_Result!jumlah_on))
        Result_SP = Trim(IIf(IsNull(M_Objrs_Result!jumlah_sp), 0, M_Objrs_Result!jumlah_sp))
        Result_POP = Trim(IIf(IsNull(M_Objrs_Result!jumlah_pop), 0, M_Objrs_Result!jumlah_pop))
        Result_PR = Trim(IIf(IsNull(M_Objrs_Result!jumlah_pr), 0, M_Objrs_Result!jumlah_pr))
        Result_PO = Trim(IIf(IsNull(M_Objrs_Result!jumlah_po), 0, M_Objrs_Result!jumlah_po))
        Result_PTP_PO = Trim(IIf(IsNull(M_Objrs_Result!jumlah_ptp_po), 0, M_Objrs_Result!jumlah_ptp_po))
        
        M_RPTCONN.execute "UPDATE TrackingRptPerPrgBatch set result_sk = '" & Result_SK & "', result_vl = '" & Result_VL & "', result_on = '" & Result_ON & "', result_sp = '" & Result_SP & "', result_pop = '" & Result_POP & "', result_pr = '" & Result_PR & "', result_po = '" & Result_PO & "', result_ptp = '" & Result_PTP_PO & "'  where AOC = '" & Descol & "' "

        
        M_Objrs_Result.MoveNext
    Wend
    
    
    
End Sub



Private Sub ReportPTPNego()
Dim Rsptp As ADODB.Recordset
Dim m_msgbox As Variant

On Error GoTo eddder:
Set Rsptp = New ADODB.Recordset
Rsptp.CursorLocation = adUseClient
If Option1(0).Value Then
cmdsql = "select agent,f_cek_NEW,count(agent) as JML,sum(promisepay) as VOL from reportPTP_new  where "
cmdsql = cmdsql + " agent in (select userid from usertbl where userid >='" + Combo2(0).text + "' and userid<='" + Combo2(1).text + "') and "
cmdsql = cmdsql + " RECSOURCE Between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and inputdate >=  "
'CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
cmdsql = cmdsql + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "'"
cmdsql = cmdsql + "  group by agent, f_cek_NEW"

Else
If Option1(1).Value Then
cmdsql = "select agent,f_cek_NEW, count(agent) as JML,sum(promisepay) as VOL from reportPTP_new  where "
cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >='" + Combo3(0).text + "' and SPVCODE<='" + Combo3(1).text + "') and "
cmdsql = cmdsql + " RECSOURCE Between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and inputdate >=  "
'CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
cmdsql = cmdsql + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "'"
cmdsql = cmdsql + "  group by agent,f_cek_NEW"
End If
End If

Rsptp.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not Rsptp.EOF
LAgent = Trim(IIf(IsNull(Rsptp!agent), "", Rsptp!agent))
jml = Trim(IIf(IsNull(Rsptp!jml), 0, Rsptp!jml))
Lf_cek = Trim(IIf(IsNull(Rsptp!f_cek_new), "", Rsptp!f_cek_new))
Lvol = Trim(IIf(IsNull(Rsptp!vol), 0, Rsptp!vol))
If Lf_cek = "PTP" Then
jml = jml
Lvol = Lvol
Else
jml = 0
Lvol = 0
End If
M_RPTCONN.execute "UPDATE TrackingRptPerPrgBatch set PTP_BARU =" + jml + ",VolPTP_Baru=" + Lvol + "  where AOC = '" + LAgent + "'"
Rsptp.MoveNext
Wend
Set Rsptp = Nothing
cmdsql = Empty
LAgent = Empty
jml = Empty
Lf_cek = Empty

Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next

End Sub
Private Sub ReportPTPNego_Compare()
Dim Rsptp As ADODB.Recordset
Dim m_msgbox As Variant

On Error GoTo eddder:
Set Rsptp = New ADODB.Recordset
Rsptp.CursorLocation = adUseClient
If Option1(0).Value Then
cmdsql = "select agent,f_cek,count(agent) as JML,sum(promisepay) as VOL from reportPTP  where "
cmdsql = cmdsql + " agent in (select userid from usertbl where userid >='" + Combo2(0).text + "' and userid<='" + Combo2(1).text + "') and "
cmdsql = cmdsql + " RECSOURCE Between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and promisedate between  "
cmdsql = cmdsql + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "' "
cmdsql = cmdsql + "  group by agent, f_cek "

Else
If Option1(1).Value Then
cmdsql = "select agent,f_cek, count(agent) as JML,sum(promisepay) as VOL from reportPTP  where "
cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >='" + Combo3(0).text + "' and SPVCODE<='" + Combo3(1).text + "') and "
cmdsql = cmdsql + " RECSOURCE Between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and promisedate between  "
cmdsql = cmdsql + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "' "
cmdsql = cmdsql + "  group by agent,f_cek "
End If
End If

Rsptp.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not Rsptp.EOF
LAgent = Trim(IIf(IsNull(Rsptp!agent), "", Rsptp!agent))
jml = Trim(IIf(IsNull(Rsptp!jml), 0, Rsptp!jml))
Lf_cek = Trim(IIf(IsNull(Rsptp!F_CEK), "", Rsptp!F_CEK))
Lvol = Trim(IIf(IsNull(Rsptp!vol), 0, Rsptp!vol))
If Lf_cek = "PTP" Then
jml = jml
Else
jml = 0
End If
M_RPTCONN.execute "UPDATE TrackingRptPerPrgBatch set PTP_BARU_LM =" + jml + ",VolPTP_Baru_LM=" + Lvol + "  where AOC = '" + LAgent + "'"
Rsptp.MoveNext
Wend
Set Rsptp = Nothing
cmdsql = Empty
LAgent = Empty
jml = Empty
Lf_cek = Empty

Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next

End Sub

Private Sub AmbilDtYgDiFU_PerAgent()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim m_msgbox As Variant
'On Error GoTo eddder

    tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
    tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
    m_hst.CursorLocation = adUseClient
    If Option1(0).Value Then
    
    'CMDSQL = "SELECT AGENT, F_CEK, COUNT(F_CEK) AS Jml,sum(ttlptp) as jmlPTP FROM"
    'CMDSQL = CMDSQL + " (select custid, recsource,F_CEK, agent,ttlptp from mgm"
    'CMDSQL = CMDSQL + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
    ''CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
    'CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    'CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
    'CMDSQL = CMDSQL + " AND (substring(F_CEK,1,2) IN ('NK','MV','WN','NA','SP','RP','BP','OP','ST')"
    'CMDSQL = CMDSQL + " or substring(F_CEK,1,3) in ('NBP','PTP','POP'))"
    'CMDSQL = CMDSQL + " AND custid in (Select distinct custid from mgm_hst"
    'CMDSQL = CMDSQL + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
    'CMDSQL = CMDSQL + " A GROUP BY AGENT,f_cek"
    
        cmdsql = "SELECT AGENT, F_CEK_NEW, COUNT(F_CEK_NEW) AS Jml FROM"
        cmdsql = cmdsql + " (select custid, recsource,F_CEK_NEW, agent,ttlptp from mgm Where "
        cmdsql = cmdsql + " RECSOURCE Between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and "
        cmdsql = cmdsql + " agent between  '" + Combo2(0).text + "' and '" + Combo2(1).text + "' AND"
        cmdsql = cmdsql + " (substring(F_CEK_NEW,1,3)"
        cmdsql = cmdsql + " in ('PO-','CO-','NBP','PTP','POP','NK-','MV-','WN-','NA-','SP-','RP-','BP-','OP-','ST-','VL-','PR-','OS-','ON-','SK-'))"
        cmdsql = cmdsql + " AND custid in (Select distinct custid from mgm_hst"
        cmdsql = cmdsql + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "')) A GROUP BY AGENT,f_cek_NEW"
    
    
    
    Else
        If Option1(1).Value Then
            cmdsql = "SELECT AGENT, F_CEK_NEW, COUNT(F_CEK_NEW) AS Jml,sum(ttlptp) as jmlPTP FROM"
            cmdsql = cmdsql + " (select custid, recsource,F_CEK_NEW, agent,ttlptp from mgm"
            cmdsql = cmdsql + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
            'CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
            cmdsql = cmdsql + " and recsource BETWEEN '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
            'CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
            cmdsql = cmdsql + " AND substring(F_CEK_NEW,1,3) IN ('PO-','CO-','NK-','MV-','WN-','SP-','NA-','RP-','BP-','OP-','ST-','NBP','PTP','POP','VL-','PR-','OS-','ON-','SK-')"
            cmdsql = cmdsql + " And agent in (select userid from usertbl where "
            cmdsql = cmdsql + " SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')"
            cmdsql = cmdsql + " AND custid in (Select distinct custid from mgm_hst "
            cmdsql = cmdsql + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
            cmdsql = cmdsql + " A GROUP BY AGENT,f_cek_NEW"
    
        End If
    End If
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        Select Case UCase(Left(IIf(IsNull(m_hst!f_cek_new), "", m_hst!f_cek_new), 3))
            Case "PTP"
                STATUS = Left(IIf(IsNull(m_hst!f_cek_new), "", m_hst!f_cek_new), 6)
            Case Else
                STATUS = Left(IIf(IsNull(m_hst!f_cek_new), "", m_hst!f_cek_new), 3)
            
            
        End Select
        cmdsql = cmdsql + "[" + STATUS + "]"
        cmdsql = cmdsql + "=[" + STATUS + "] + " + CStr(IIf(IsNull(m_hst!jml), 0, m_hst!jml)) + " "
        If Left(STATUS, 3) = "PTP" Then
        'CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!f_cek_new) Then
        Else
            If m_hst!f_cek_new = "" Then
            Else
            
                If m_hst!jml = 0 Then
                Else
                    
                    M_RPTCONN.execute cmdsql
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
LAgent = Empty
cmdsql = Empty
Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    'Resume Next
End Sub

Private Sub AmbilDtYgDiFU_PerAgentcall()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim m_msgbox As Variant
On Error GoTo eddder

tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
If Option1(0).Value Then
cmdsql = " SELECT AGENT, StatusCall, COUNT(StatusCall) AS Jml,sum(ttlptp) as jmlPTP FROM"
cmdsql = cmdsql + " (select custid, recsource,StatusCall, agent,ttlptp from mgm"
cmdsql = cmdsql + " where tglcall >= '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and tglcall <= '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'"
cmdsql = cmdsql + " and recsource >= '" + Combo1(0).text + "' and recsource <= '" + Combo1(1).text + "'"
'CMDSQL = CMDSQL + " And Kethslkerja <> 'I') "
cmdsql = cmdsql + " And custid in (Select distinct custid from mgm_hst"
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
cmdsql = cmdsql + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
cmdsql = cmdsql + " A GROUP BY AGENT, STATUSCALL"

'CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
'CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
'CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and "
'CMDSQL = CMDSQL + " custid in(select custid from mgm where agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ) "
'CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
'CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
'CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
'CMDSQL = CMDSQL + "AND mgm_hst.F_CEK IN ('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
'CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
'CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
'CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
'CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and "
'CMDSQL = CMDSQL + " custid in(select custid from mgm where agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ) "
'CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
'CMDSQL = CMDSQL + " inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
'CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
'CMDSQL = CMDSQL + " AND (left(mgm_hst.F_CEK,2) IN ('NK','MV','WN','NA','RP','BP','OP','ST') or left(mgm_hst.F_CEK,3) in ('NBP','PTP','POP')) "
'CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"

Else
If Option1(1).Value Then
cmdsql = " SELECT AGENT, StatusCall, COUNT(StatusCall) AS Jml,sum(ttlptp) as jmlPTP FROM"
cmdsql = cmdsql + " (select custid, recsource,StatusCall, agent,ttlptp from mgm"
cmdsql = cmdsql + " where tglcall >= '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and tglcall <= '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'"
cmdsql = cmdsql + " and recsource >= '" + Combo1(0).text + "' and recsource <= '" + Combo1(1).text + "'"
'CMDSQL = CMDSQL + " And Kethslkerja <> 'I') "
cmdsql = cmdsql + " And agent in (select userid from usertbl where "
cmdsql = cmdsql + " SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')"
cmdsql = cmdsql + " And custid in (Select distinct custid from mgm_hst"
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
cmdsql = cmdsql + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
cmdsql = cmdsql + " A GROUP BY AGENT, STATUSCALL"


'CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
'CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
'CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and "
'CMDSQL = CMDSQL + " custid in(select custid from mgm where agent in (select userid from usertbl where "
'CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "'))"
'CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
'CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
'CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
'CMDSQL = CMDSQL + "AND mgm_hst.F_CEK IN ('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
'CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent "

'CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
'CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
'CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and "
'CMDSQL = CMDSQL + " custid in(select custid from mgm where agent in (select userid from usertbl where "
'CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "'))"
'CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
'CMDSQL = CMDSQL + " inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
'CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
'CMDSQL = CMDSQL + " AND (left(mgm_hst.F_CEK,2) IN ('NK','MV','WN','NA','RP','BP','OP','ST') or left(mgm_hst.F_CEK,3) in ('NBP','PTP','POP')) "
'CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent "
End If
End If
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        Select Case Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 2)
            Case "NK"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3) & "HO"
            Case "MV"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3) & "HO"
            Case "WN"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3) & "HO"
            Case "BP"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3)
            Case "PT"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3)
            Case "RP"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3) & "B"
            Case "NA"
                'If Len(m_hst!StatusCall) = 5 Then
            STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3) & "O"
                'ElseIf Len(m_hst!StatusCall) = 4 Then
                    'STATUS = Left(IIf(IsNull(m_hst!StatusCall), "", m_hst!StatusCall), 4)
                'End If
            Case "SP"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3)
              Case "PO"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3)
             Case "OP"
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3)
            Case "NB"
                STATUS = Trim(Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 3))
            Case Else
                STATUS = Left(IIf(IsNull(m_hst!statuscall), "", m_hst!statuscall), 4)
            
        End Select
        cmdsql = cmdsql + "[" + STATUS + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!jml), 0, m_hst!jml)) + " "
        If STATUS = "PTP" Then
        cmdsql = cmdsql + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!statuscall) Then
        Else
            If m_hst!statuscall = "" Then
            Else
            
                If m_hst!jml = 0 Then
                Else
                    M_RPTCONN.execute cmdsql
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
LAgent = Empty
cmdsql = Empty
Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next
End Sub

Private Sub AmbilDtYgDiFU_PerAgent_Compare()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim m_msgbox As Variant
On Error GoTo eddder

tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
If Option1(0).Value Then
cmdsql = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
cmdsql = cmdsql + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
cmdsql = cmdsql + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "'  and "
cmdsql = cmdsql + " custid in(select custid from mgm where agent between  '" + Combo2(0).text + "' and '" + Combo2(1).text + "' ) "
cmdsql = cmdsql + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
cmdsql = cmdsql + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
cmdsql = cmdsql + " recsource BETWEEN '" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
cmdsql = cmdsql + "AND mgm_hst.F_CEK IN ('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
cmdsql = cmdsql + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
Else
If Option1(1).Value Then
cmdsql = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
cmdsql = cmdsql + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
cmdsql = cmdsql + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "'  and "
cmdsql = cmdsql + " custid in(select custid from mgm where agent in (select userid from usertbl where "
cmdsql = cmdsql + " SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "'))"
cmdsql = cmdsql + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
cmdsql = cmdsql + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
cmdsql = cmdsql + " recsource BETWEEN '" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
cmdsql = cmdsql + "AND mgm_hst.F_CEK IN ('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
cmdsql = cmdsql + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent "
End If
End If
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        Select Case Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 2)
            Case "NK"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "MV"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "WN"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "BP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "PT"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "RP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            Case "NA"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            Case "SP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
              Case "PO"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
             Case "OP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "NB"
                STATUS = Trim(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3))
            Case Else
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            
        End Select
        Last = "_LM"
        STATUS = STATUS + Last
        cmdsql = cmdsql + "[" + STATUS + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!jml), 0, m_hst!jml)) + " "
        If STATUS = "PTP_LM" Then
        cmdsql = cmdsql + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!F_CEK) Then
        Else
            If m_hst!F_CEK = "" Then
            Else
            
                If m_hst!jml = 0 Then
                Else
                    M_RPTCONN.execute cmdsql
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
LAgent = Empty
cmdsql = Empty
Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next
End Sub
Private Sub AmbilDtYgDiFU_PerFIELD()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim m_msgbox As Variant

On Error GoTo eddder
tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
If Option1(0).Value Then
cmdsql = "select  tblvisit.ffc, tblvisit.StatusVisit,count(tblvisit.statusVisit) as jml from tblvisit "
cmdsql = cmdsql + " inner join (SELECT custid, ffc, max(RequestDate) as tglmax from tblvisit where "
cmdsql = cmdsql + " RequestDate Between  '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
cmdsql = cmdsql + " and ffc in (select userid from usertbl where userid >= '" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' )group by custid,ffc)  as  a "
cmdsql = cmdsql + " on tblvisit.custid = a.custid and tblvisit.requestdate=a.tglmax "
cmdsql = cmdsql + " inner join mgm on mgm.custid= a.custid where tblvisit.statusvisit in('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
cmdsql = cmdsql + " and recsource BETWEEN '" + Combo1(0).text + "' and '" + Combo1(1).text + "'  "
cmdsql = cmdsql + "group by tblvisit.ffc, statusVisit "
Else
If Option1(1).Value Then
cmdsql = "select  tblvisit.ffc, tblvisit.StatusVisit,count(tblvisit.statusVisit) as jml from tblvisit "
cmdsql = cmdsql + " inner join (SELECT custid, ffc, max(RequestDate) as tglmax from tblvisit where "
cmdsql = cmdsql + " RequestDate Between  '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
cmdsql = cmdsql + " and ffc in (select userid from usertbl where SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "' )group by custid,ffc)  as  a "
cmdsql = cmdsql + " on tblvisit.custid = a.custid and tblvisit.requestdate=a.tglmax "
cmdsql = cmdsql + " inner join mgm on mgm.custid= a.custid where tblvisit.statusvisit in('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
cmdsql = cmdsql + " and recsource BETWEEN '" + Combo1(0).text + "' and '" + Combo1(1).text + "'  "
cmdsql = cmdsql + "group by tblvisit.ffc, statusVisit "
End If
End If
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!FFC), "", m_hst!FFC)))
        cmdsql = "Update TrackingRptField Set "
        Select Case Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 2)
            Case "NK"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 5)
            Case "MV"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 5)
            Case "WN"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 5)
            Case "BP"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 3)
            Case "PT"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 3)
            Case "RP"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 4)
            Case "NA"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 4)
            Case "SP"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 3)
              Case "PO"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 3)
             Case "OP"
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 3)
            Case "NB"
                STATUS = Trim(Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 3))
            Case Else
                STATUS = Left(IIf(IsNull(m_hst!StatusVisit), "", m_hst!StatusVisit), 4)
            
        End Select
        cmdsql = cmdsql + "[" + STATUS + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!jml), 0, m_hst!jml)) + " "
'        If STATUS = "PTP" Then
'        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
'        End If
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!StatusVisit) Then
        Else
            If m_hst!StatusVisit = "" Then
            Else
            
                If m_hst!jml = 0 Then
                Else
                    M_RPTCONN.execute cmdsql
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
cmdsql = Empty
LAgent = Empty
Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next
End Sub


Private Sub AmbilDataYgDiFU_LastMonth()
Dim m_hst As New ADODB.Recordset
Dim LastMonth As String
Dim m_msgbox As Variant
Dim VarMonth As String
Dim VarYear As String


On Error GoTo eddder
VarYear = Format(TDBDate1(0).Value, "yyyy")
VarMonth = Format(TDBDate1(1).Value, "mm")

If VarMonth = 1 Then
VarMonth = "12"
VarYear = (VarYear) - 1
Else
VarMonth = (VarMonth) - 1
VarYear = VarYear
End If

m_hst.CursorLocation = adUseClient
If Option1(0).Value Then
cmdsql = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
cmdsql = cmdsql + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where MONTH(tgl) ='" + VarMonth + "' AND YEAR(tgl)='" + VarYear + "' AND "
cmdsql = cmdsql + " custid in(select custid from mgm where agent between  '" + Combo2(0).text + "' and '" + Combo2(1).text + "' ) "
cmdsql = cmdsql + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
cmdsql = cmdsql + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
cmdsql = cmdsql + " recsource BETWEEN '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
cmdsql = cmdsql + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
Else
If Option1(1).Value Then
cmdsql = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
cmdsql = cmdsql + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where MONTH(tgl) ='" + VarMonth + "' AND YEAR(tgl)='" + VarYear + "' AND "
cmdsql = cmdsql + " custid in(select custid from mgm where agent in (select userid from usertbl where "
cmdsql = cmdsql + " SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "'))"
cmdsql = cmdsql + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
cmdsql = cmdsql + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
cmdsql = cmdsql + " recsource BETWEEN '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
cmdsql = cmdsql + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
 End If
End If
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        Select Case Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 2)
            Case "NK"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "MV"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "WN"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "BP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "PT"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "RP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            Case "NA"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            Case "SP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
              Case "PO"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
             Case "OP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case Else
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
             End Select
         LastMonth = "_LM"
        cmdsql = cmdsql + "[" + STATUS + LastMonth + "]"
'        CMDSQL = CMDSQL + "[" + STATUS + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!jml), 0, m_hst!jml)) + " "
        If STATUS = "PTP" Then
        cmdsql = cmdsql + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!F_CEK) Then
        Else
            If m_hst!F_CEK = "" Then
            Else
                If m_hst!jml = 0 Then
                Else
                   
                    M_RPTCONN.execute cmdsql
                End If
            End If
        End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing

Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
  '      MsgBox Err.Description
    End If
    Resume Next

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
    Check3.Value = 0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
    Check3.Value = 0
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Check1.Value = 0
    Check2.Value = 0
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
    Call Combo1_LostFocus(Index)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_Objrs As New ADODB.Recordset
On Error GoTo Combo1_LostFocusErr
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open "Select * from datasourcetbl where kodeds ='" + Combo1(Index).text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not M_Objrs.EOF Then
        Combo1(Index).text = M_Objrs!KODEDS
    Else
        Combo1(Index).text = Empty
    End If
Exit Sub
Combo1_LostFocusErr:
    MsgBox err.Description
End Sub

Private Sub Combo2_Click(Index As Integer)
    Call Combo2_LostFocus(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim M_Objrs As New ADODB.Recordset
On Error GoTo Combo2_LostFocusErr
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND USERID ='" + Combo2(Index).text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not M_Objrs.EOF Then
        Combo2(Index).text = M_Objrs!Userid
    Else
        Combo2(Index).text = Empty
    End If
Exit Sub
Combo2_LostFocusErr:
    MsgBox err.Description
End Sub
Private Sub hitung_JmlData_PerAgent_PTP()
Dim M_Objrs As New ADODB.Recordset
Dim ptp As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_Objrs.CursorLocation = adUseClient

'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
cmdsql = "Select Agent, sum(ttlptp) as JMLVOL, count(f_cek_NEW) as PTP from mgm  where F_CEK_NEW='PTP' AND recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' AND Tglincoming BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    'JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "0", m_objrs!jml))
    JUMLAHVOL = CStr(IIf(IsNull(M_Objrs!JMLVOL), "0", M_Objrs!JMLVOL))
    LAgent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    ptp = CStr(IIf(IsNull(M_Objrs!ptp), "0", M_Objrs!ptp))
    cmdsql = "Update TrackingRptPerPrgBatch set  VOLPTP1 = " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.execute cmdsql
    cmdsql = "Update TrackingRptPerPrgBatch set  PTP1 = " + ptp + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.execute cmdsql
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
JUMLAHVOL = Empty
LAgent = Empty
ptp = Empty
cmdsql = Empty


Exit Sub
hitung_JmlDataer:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim rsTemp1 As New ADODB.Recordset
'On Error GoTo Command1_ClickeR
Dim SYARAT As String
Dim Strsql As String
If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
    TDBDate1(0).Value = "01/01/1990"
    TDBDate1(1).Value = "31/12/2020"
End If

If Combo1(0).text = Empty And Combo1(1).text = Empty Then
    Combo1(0).text = "-----"
    Combo1(1).text = "ZZZZZ"
End If
If Option1(0).Value = False And Option1(1).Value = False Then
If Combo2(0).text = Empty And Combo2(1).text = Empty Then
    Combo2(0).text = "-----"
    Combo2(1).text = "ZZZZZ"
End If
End If

            
            
ProgressBar1.Visible = True

'Tambahan dari Izuddin Buat Excel
b_excel = False
Select Case Index
    Case 2
        Command1(2).Enabled = False
        ' Tarik data ke excel
        b_excel = True
        Select Case ListView1.SelectedItem.text
        Case 39
            Call RptRequestPTP
        Case 38
            Call TarikSeluruhCPA
        'Randy 26March2015
        Case 41
            Call Isi_Report_PTP_REG_Jatuh_Tempo_Excel
        End Select
        ' ------------------
        Command1(2).Enabled = True
    
    Case 0
    
    Select Case ListView1.SelectedItem.text
        
              
        Case 40 ' @@ 19-04-2013
            Call RptDetailPayment_Interval_permonth
              
        Case 39 '@@27-02-2013 Report Request PTP
            Call RptRequestPTP
            WaitSecs (2)
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Rptlist.rpt"
            Call SHOW_PRN
        
        Case 38 '@@Report untuk menarik seluruh CPA
            Call TarikSeluruhCPA
            Call UpdateAllPaymentCPA
            WaitSecs (2)
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptALLCPA_NEW.rpt"
            Call SHOW_PRN
            
            
        Case 37 '@@20072012 Report Durasi Call Contact LPD Server 4
            Call HitungDurasiLPDIcentra5
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptDurasiCallLPD.rpt"
            Call SHOW_PRN
         
         Case 36 '@@20072012 Report Durasi Call Contact LPD Server 4
            Call HitungDurasiLPDIcentra4
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptDurasiCallLPD.rpt"
            Call SHOW_PRN
         
         Case 35 '@@13072012 Report CPA terakhir
            Call IsiCustidPaidOff
            WaitSecs (2)
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPaidOff.rpt"
            Call SHOW_PRN
         
         Case 34 '@@09072012 ContactRate dan ContactLPD1
            Call ContactRateLPD1
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptContactRateLPD.rpt"
            Call SHOW_PRN
         
         '@@21062012 Report TblSendPTP Approve
         Case 32
            Call LogApprovalSendPTP
            WaitSecs (2)
            RPT.Reset
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptSendPTPApproved.rpt"
            Call SHOW_PRN
         Case 33
            Call LogRejectedSendPTP
            WaitSecs (2)
            RPT.Reset
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptSendPTPRejected.rpt"
            Call SHOW_PRN
         
         '@@14 May 2012 Bikin Report UnValid Number
         Case 29 'Report UnValid Number
            Call Isi_Unvalid_Number
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptUnValidNumber.rpt"
            Call SHOW_PRN
            
         Case 30 'Report Valid Number
            Call Isi_Valid_Number
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptValidNumber.rpt"
            Call SHOW_PRN
         Case 31 'Report Review Account
            Call RptAccReview
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptReviewAcc.rpt"
            Call SHOW_PRN
             
         
         '@@ 05 Mei 2011 Report Request BS
         Case 21
            Call Isi_Bs
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptReqBS.rpt"
            Call SHOW_PRN
         Case 22
            Call Isi_EC
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptReqEC.rpt"
            Call SHOW_PRN
         Case 23
            Call Isi_OST
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptReqOST.rpt"
            Call SHOW_PRN
         Case 24
            Call Isi_Problem
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptReqProblem.rpt"
            Call SHOW_PRN
         Case 25
            Call Isi_PUM
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptReqPUM.rpt"
            Call SHOW_PRN
        Case 26
            Call Isi_RS
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptReqRS.rpt"
            Call SHOW_PRN
         
'        Case 1
'            Call Isi_Agent_mgm
'            Call hitung_JmlData_PerAgent_mgm
'            Call AmbilDtYgDiFU_PerAgent
'            Call VolumeUtilized
'            Call ReportPTPNego
'            Call Isi_Settled_Payment
'            Call hitung_BatchCallInitilized_PerAgent_mgm
'            Call Hitung_Number_of_Payment
'            Call Hitung_Volume_of_Payment
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\Tracking ReportSPV_sum.rpt"
'            Call SHOW_PRN
'
'         Case 2
'            Call Isi_Agent_mgm
'            Call hitung_JmlData_PerAgent_mgm
'            Call AmbilDtYgDiFU_PerAgent
'            Call VolumeUtilized
'            Call ReportPTPNego
'            Call Isi_Settled_Payment
''           Call Isi_Progess_OF_PAyment
'            Call hitung_JmlData_PerAgent_PTP
'          '' Call Hitung_JmlLeadsPerAgent
'          Call Hitung_Vol_PTP
'          Call hitung_BatchCallInitilized_PerAgent_mgm
'          Call Hitung_Number_of_Payment
'          Call Hitung_Volume_of_Payment
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\Tracking ReportAgent.rpt"
'
'            Call SHOW_PRN
        
       Case 1
            Call Isi_Agent_mgm
            Call hitung_JmlData_PerAgent_mgm
            Call AmbilDtYgDiFU_PerAgent
            Call VolumeUtilized
            Call ReportPTPNego
            'Call AmbilDataYgDiFU_LastMonth
            'Call Isi_Settled_Payment
'           Call Isi_Progess_OF_PAyment
            Call hitung_JmlData_PerAgent_PTP
          ' Call Hitung_JmlLeadsPerAgent
           Call Hitung_Vol_PTP
           Call hitung_BatchCallInitilized_PerAgent_mgm
           'Call Hitung_Number_of_Payment
           'Call Hitung_Volume_of_Payment
           'Call Hitung_Payment
           
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Tracking ReportSPVGlobalNew.rpt"
            Call SHOW_PRN
        
        Case 2
            Call Isi_Agent_mgm
            Call Hitung_Pergerakan_status
            Call hitung_JmlData_PerAgent_mgm
            Call AmbilDtYgDiFU_PerAgent
            Call VolumeUtilized
            Call ReportPTPNego
            'Call AmbilDataYgDiFU_LastMonth
          '  Call Isi_Settled_Payment
'            Call Isi_Progess_OF_PAyment
            Call hitung_JmlData_PerAgent_PTP
          ' Call Hitung_JmlLeadsPerAgent
            Call Hitung_Vol_PTP
            Call hitung_BatchCallInitilized_PerAgent_mgm
'            Call Hitung_Number_of_Payment
          '  Call Hitung_Volume_of_Payment
        
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Tracking ReportagentGlobalNewBR.rpt"
            Call SHOW_PRN
            
'        Case 19
'            Call Isi_Agent_mgm
'            Call hitung_JmlData_PerAgent_mgm
'            Call AmbilDtYgDiFU_PerAgentcall
'            Call VolumeUtilized
'            Call hitung_BatchCallInitilized_PerAgent_mgm
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportAgentGlobalstatuscallNew.rpt"
'            Call SHOW_PRN
            
        Case 3
            
            Call Isi_Data_Distribusi
            WaitSecs (2)
            RPT.Reset
'            If TDBDate1(0).ValueIsNull = False And TDBDate1(1).ValueIsNull = False Then
'            cmdsql = "{mgm.tglsource} IN DATE(#" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & "#) " & _
'                          "TO DATE (#" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & "#)"
'            End If
            
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@SPV1 = totext('" + CStr(Combo3(0).Text) + "')"
'            RPT.Formulas(3) = "@SPV2 = totext('" + CStr(Combo3(1).Text) + "')"
'            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
'            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            'RPT.SelectionFormula = cmdsql
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptDistribusi.rpt"
            Call SHOW_PRN
         
'        Case 6
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
'            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
'            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
'            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
'            RPT.ReportFileName = App.Path + "\Report\historyCall.rpt"
'            Call SHOW_PRN
         
'        Case 7
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
'            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
'            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
'            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
'            RPT.ReportFileName = App.Path + "\Report\historyCall_custid.rpt"
'            Call SHOW_PRN
'
                 
'        Case 8
'        Call Tracking_PTP_JatuhTempo_NEW
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
'            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
'            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
'            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
'            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
'            'RPT.ReportFileName = "D:\COLLECTION\Report\ReservedPTP.rpt"
'            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew.rpt"
'            Call SHOW_PRN
         
         Case 4
         Call Isi_Report_PTP_lunas
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ActualPay.rpt"
            Call SHOW_PRN
            
        '@@ 04-05-2011 Report PTP Jatuh Tempo
        Case 5
           Call Isi_Report_PTP_Jatuh_Tempo
'            Tracking_PTP_JatuhTempo_NEW
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew.rpt"
            Call SHOW_PRN
            
        'Randy 26March2015
        Case 41
           Call Isi_Report_PTP_REG_Jatuh_Tempo
'            Tracking_PTP_JatuhTempo_NEW
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew_REGULER.rpt"
            Call SHOW_PRN
        
        'Randy 26March2015
        Case 42
           Call Isi_Report_PTP_PO_Jatuh_Tempo
'            Tracking_PTP_JatuhTempo_NEW
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew.rpt"
            Call SHOW_PRN
                
        Case 43
           Call Isi_Report_PTP_REGULER2
'            Tracking_PTP_JatuhTempo_NEW
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew_REGULER2.rpt"
            Call SHOW_PRN
         
        Case 44
           Call Isi_Report_On_Going_PTP
'            Tracking_PTP_JatuhTempo_NEW
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew_REGULER2.rpt"
            Call SHOW_PRN
            
        'DODDY REQUEST 18MEI2015(Randy)
        Case 45
           Call Isi_Report_Result_DeskCall
'            Tracking_PTP_JatuhTempo_NEW
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptResultDeskCall.rpt"
            Call SHOW_PRN
         
         
         Case 6
     '    Call GET_PTP_NEW
            Call GETPTPNEW
'         Call Isi_Report_FormVisit
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPromiseToPaynew.rpt"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptVisit.rpt"
            Call SHOW_PRN
        
           Case 7
           
            Call TrackingReservedPTP
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReservedPTP.rpt"
            Call SHOW_PRN
            
             Case 16
           
            Call TrackingReservedPTPversi2
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReservedPTPversi2.rpt"
            Call SHOW_PRN
            
        '@@Report POP-BP1 [24-11-09] -- POSTGREE
        Case 8
            Call ISI_DATA_POP_BP1
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPOPBP.rpt"
            Call SHOW_PRN
            
            
            
        '@@Report POP-BP2 [24-11-09] -- POSTGREE
        Case 9
            Call ISI_DATA_POP_BP2
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPOPBP2.rpt"
            Call SHOW_PRN
        
        '@@Report POP-BP3 [24-11-09] -- POSTGREE
        Case 10
            Call ISI_DATA_POP_BP3
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPOPBP3.rpt"
            Call SHOW_PRN
             
        '@@Report BP1 [24-11-09] -- POSTGREE
        Case 11
            Call ISI_DATA_BP1
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptBP1.rpt"
            Call SHOW_PRN
        
        '@@Report BP2 [24-11-09] -- POSTGREE
        Case 12
            Call isidataBpMonth
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptBpmonth.rpt"
            Call SHOW_PRN
        
        '@@Report BP3 [24-11-09] -- POSTGREE
        Case 13
              Call isidatabpday2
           ' Call isidatabpday
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptBpday.rpt"
            Call SHOW_PRN
        
        Case 14
            Call PTP_log
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNewlog.rpt"
            Call SHOW_PRN
        
        Case 15
            Call CariCPAApprove
            Call CariCPAPaidOff
            Call CariCPARejected
            Call CariCPA_ToBe_Approve_Hamanto
            WaitSecs (2)
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Rptlist.rpt"
            Call SHOW_PRN

            
         Case 17
           RPT.Reset
            If TxtCustid.text <> "" Then
                     If Len(SYARAT) > 0 Then
                    SYARAT = SYARAT + " AND vcustid ='" + TxtCustid.text + "'"
                    Else
                    SYARAT = " WHERE  vcustid ='" + TxtCustid.text + "'"
                End If
            End If
            
            If Len(SYARAT) > 0 Then
                    SYARAT = SYARAT + " AND dtglinsert between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
                    Else
                    SYARAT = " WHERE  dtglinsert  between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
                End If
                
                
           M_RPTCONN.execute "delete from tblreportcpa "
           Strsql = "select * from tblreportcpa"
           Set rsTemp1 = New ADODB.Recordset
           rsTemp1.CursorLocation = adUseClient
           rsTemp1.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
           
           
           'CMDSQL = " SELECT * FROM (SELECT *  FROM USERTBL) AS B"
           'CMDSQL = CMDSQL + " Right JOIN  ("
           'CMDSQL = CMDSQL + " SELECT * FROM ("
           'CMDSQL = CMDSQL + " SELECT  * FROM (  SELECT * FROM TBLCPA ) AS A Inner Join"
           'CMDSQL = CMDSQL + " (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID " + SYARAT + " ))  AS BRU ON BRU.AGENT=B.USERID"

           cmdsql = "  SELECT * FROM ( "
          cmdsql = cmdsql + " SELECT * FROM (SELECT *  FROM USERTBL) AS B"
          cmdsql = cmdsql + " Right JOIN  ( "
          cmdsql = cmdsql + " SELECT * FROM ( "
          cmdsql = cmdsql + " SELECT  * FROM (  SELECT * FROM TBLCPA ) AS A Inner Join "
          cmdsql = cmdsql + "  (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID " + SYARAT + "   ))  AS BRU ON BRU.AGENT=B.USERID) AS TBLBARU"
          cmdsql = cmdsql + " left Join ( "
          cmdsql = cmdsql + "   select * from ( "
          cmdsql = cmdsql + " SELECT custid as cust_no,PAYDATE AS lpd,payment as lpa FROM TBLLUNAS  WHERE ID IN (SELECT MAX(ID) FROM tbllunas GROUP BY CUSTID))  as tblbaru1 ) as bru on tblbaru.custid=bru.cust_no "
          


           
           
'           cmdsql = "SELECT  * FROM ( "
'           cmdsql = cmdsql + " SELECT * FROM TBLCPA ) AS A"
'           cmdsql = cmdsql + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID " + SYARAT


           Set rsTemporary = New ADODB.Recordset
           rsTemporary.CursorLocation = adUseClient
          
           rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
           ProgressBar1.Max = rsTemporary.RecordCount + 1
           While Not rsTemporary.EOF
               ProgressBar1.Value = rsTemporary.Bookmark
            DoEvents
            rsTemp1.AddNew
            rsTemp1("dtglinsert") = IIf(IsNull(rsTemporary("dtglinsert")), "", rsTemporary("dtglinsert"))
            rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
            rsTemp1("dproposal") = IIf(IsNull(rsTemporary("dpropsal")), Null, rsTemporary("dpropsal"))
            rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
            rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
            rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
            rsTemp1("cardno") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
            rsTemp1("custname") = IIf(IsNull(rsTemporary("name")), "", rsTemporary("name"))
            rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
            rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
            rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
            rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
            rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
            rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), "", rsTemporary("nfuturepay"))
            rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
            rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
            rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
            rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
            rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
            rsTemp1("custid") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
            rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
            rsTemp1("vjust") = IIf(IsNull(rsTemporary("vjust")), "", rsTemporary("vjust"))
            rsTemp1("agency") = IIf(IsNull(rsTemporary("agency")), "", rsTemporary("agency"))
            rsTemp1("vnameverify") = IIf(IsNull(rsTemporary("vnameverify")), "", rsTemporary("vnameverify"))
            rsTemp1("vreason") = IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason"))
            rsTemp1("vdlq") = IIf(IsNull(rsTemporary("vnodlq")), "", rsTemporary("vnodlq"))
            rsTemp1("vpaymenthandle") = IIf(IsNull(rsTemporary("vpaymenthandle")), "", rsTemporary("vpaymenthandle"))
            rsTemp1("voccupation") = IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation"))
            rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
            rsTemp1("dtglapprove") = IIf(IsNull(rsTemporary("tglapprove")), Null, rsTemporary("tglapprove"))
            rsTemp1("userid") = IIf(IsNull(rsTemporary("userid")), "", rsTemporary("userid"))
            rsTemp1("team") = IIf(IsNull(rsTemporary("team")), "", rsTemporary("team"))
            rsTemp1("chkfaxed") = IIf(IsNull(rsTemporary("chkfaxed")), "", rsTemporary("chkfaxed"))
            rsTemp1("chkwentalking") = IIf(IsNull(rsTemporary("chkwentalking")), "", rsTemporary("chkwentalking"))
            rsTemp1("chkktp") = IIf(IsNull(rsTemporary("chkktp")), "", rsTemporary("chkktp"))
            rsTemp1("chksup") = IIf(IsNull(rsTemporary("chksup")), "", rsTemporary("chksup"))
            rsTemp1("chkbillings") = IIf(IsNull(rsTemporary("chkbillings")), "", rsTemporary("chkbillings"))
            rsTemp1("chkothers") = IIf(IsNull(rsTemporary("chkothers")), "", rsTemporary("chkothers"))
            rsTemp1("ketother") = IIf(IsNull(rsTemporary("ketother")), "", rsTemporary("ketother"))
            rsTemp1("ed") = IIf(IsNull(rsTemporary("tglsource")), Null, rsTemporary("tglsource"))
            rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, rsTemporary("b_d"))
            rsTemp1("odt") = IIf(IsNull(rsTemporary("OpenDate")), Null, rsTemporary("OpenDate"))
            If IIf(IsNull(rsTemporary("lpd")), "", rsTemporary("lpd")) = "" Then
                  rsTemp1("lpd") = IIf(IsNull(rsTemporary("pay_dt")), Null, rsTemporary("pay_dt"))
            Else
                  rsTemp1("lpd") = IIf(IsNull(rsTemporary("lpd")), Null, rsTemporary("lpd"))
            End If
            
            If IIf(IsNull(rsTemporary("lpa")), "", rsTemporary("lpa")) = "" Then
                  rsTemp1("lpa") = IIf(IsNull(rsTemporary("LastPay")), 0, rsTemporary("LastPay"))
            Else
                  rsTemp1("lpa") = IIf(IsNull(rsTemporary("lpa")), 0, rsTemporary("lpa"))
            End If
            
            rsTemp1.update
           
                    rsTemporary.MoveNext
           Wend
           
          
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptCpaRincian.rpt"
            WaitSecs (2)
            Call SHOW_PRN
            Set rsTemp1 = Nothing
            Set rsTemporary = Nothing
            
            
            
'            'RPT.Reset
'            M_RPTCONN.Execute "delete from tblreportcpa "
'           strsql = "select * from tblreportcpa"
'           Set rstemp1 = New ADODB.Recordset
'           rstemp1.CursorLocation = adUseClient
'           rstemp1.Open strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
'           cmdsql = "SELECT * from " + Chr(34) + "REPORTTEST" + Chr(34) + " " + SYARAT
'          Set rsTemporary = New ADODB.Recordset
'          rsTemporary.CursorLocation = adUseClient
'
'           rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'           ProgressBar1.Max = rsTemporary.RecordCount + 1
'           While Not rsTemporary.EOF
'                ProgressBar1.Value = rsTemporary.Bookmark
'            DoEvents
'            rstemp1.AddNew
'
'            rstemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
'            rstemp1("dproposal") = IIf(IsNull(rsTemporary("dpropsal")), Null, rsTemporary("dpropsal"))
'            rstemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
'            rstemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
'            rstemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
'            rstemp1("cardno") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
'            rstemp1("custname") = IIf(IsNull(rsTemporary("name")), "", rsTemporary("name"))
'            rstemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
'            rstemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
'            rstemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
'            rstemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
'            rstemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
'            rstemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), "", rsTemporary("nfuturepay"))
'            rstemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
'            rstemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
'            rstemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
'            rstemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
'            rstemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
'            rstemp1("custid") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
'            rstemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
'            rstemp1("vjust") = IIf(IsNull(rsTemporary("vjust")), "", rsTemporary("vjust"))
'            rstemp1("agency") = IIf(IsNull(rsTemporary("agency")), "", rsTemporary("agency"))
'            rstemp1("vnameverify") = IIf(IsNull(rsTemporary("vnameverify")), "", rsTemporary("vnameverify"))
'            rstemp1("vreason") = IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason"))
'            rstemp1("vdlq") = IIf(IsNull(rsTemporary("vnodlq")), "", rsTemporary("vnodlq"))
'            rstemp1("vpaymenthandle") = IIf(IsNull(rsTemporary("vpaymenthandle")), "", rsTemporary("vpaymenthandle"))
'            rstemp1("voccupation") = IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation"))
'            rstemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
'            rstemp1("dtglapprove") = IIf(IsNull(rsTemporary("tglapprove")), Null, rsTemporary("tglapprove"))
'            rstemp1.UPDATE
'
'                    rsTemporary.MoveNext
'           Wend
'
'
'            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptCpaRincian.rpt"
'            WaitSecs (2)
'            Call SHOW_PRN
'            Set rstemp1 = Nothing
'            Set rsTemporary = Nothing
         
         
         
        
'        Case 16
'
'           If Combo2(0).Text <> "" Or Combo2(0).Text <> "" Then
'            SYARAT = "where REPORTTEST.agent between '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
'            End If
'
'            If Combo1(0).Text <> "" Or Combo1(1).Text <> "" Then
'                If Len(SYARAT) > 0 Then
'                    SYARAT = SYARAT + " AND recsource BETWEEN '" + Combo1(0).Text + "' AND '" + Combo1(1).Text + "'"
'                    Else
'                    SYARAT = " WHERE  recsource BETWEEN '" + Combo1(0).Text + "' AND '" + Combo1(1).Text + "'"
'                End If
'            End If
'
'            If Combo3(0).Text <> "" Or Combo3(1).Text <> "" Then
'                If Len(SYARAT) > 0 Then
'                    SYARAT = SYARAT + " AND REPORTTEST.agent IN ( SELECT userid FROM USERTBL WHERE SPVCODE  BETWEEN '" + Combo3(0).Text + "' AND '" + Combo3(1).Text + "')"
'                    Else
'                    SYARAT = " WHERE dtgllastupdate  REPORTTEST.agent IN ( SELECT userid FROM USERTBL WHERE SPVCODE  BETWEEN '" + Combo3(0).Text + "' AND '" + Combo3(1).Text + "')"
'                End If
'            End If
'
'            If Len(SYARAT) > 0 Then
'                    SYARAT = SYARAT + " AND dtgllastupdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
'                    Else
'                    SYARAT = " WHERE  between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
'                End If
'
'            RPT.Reset
'



'            CMDSQL = "SELECT spvcode, vregion, dpropsal,reffno,vcustid, vproduct,"
'            CMDSQL = CMDSQL + "  varragement, vcardsts,nttlpayment,ndownpay, nfuturepay,ncharge,ndiscountamt,vosbalance ,vosprincipal,"
'            CMDSQL = CMDSQL + " vjust, voccupation,vreason,vnodlq, vosbalance,vpaymenthandle ,nbalance,nprincipal,agency,name,"
'            CMDSQL = CMDSQL + " usertbl.agent ,opendate ,vnameverify From  { oj " + Chr(34) + "public" + Chr(34) + "." + Chr(34) + "usertbl" + Chr(34) + " usertbl "
'            CMDSQL = CMDSQL + "inner join " + Chr(34) + "public" + Chr(34) + "." + Chr(34) + "REPORTTEST" + Chr(34) + " REPORTTEST on "
'            CMDSQL = CMDSQL + "usertbl." + Chr(34) + "userid" + Chr(34) + "=REPORTTEST." + Chr(34) + "agent" + Chr(34) + "  } "
'            RPT.SQLQuery = CMDSQL + SYARAT
'            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptCpaRincian.rpt"
'            Call SHOW_PRN
            'Call SHOW_PRN
            
'         Case 12
'         WaitSecs (2)
'          Call Isi_Field_Collector
'          Call hitung_JmlData_FieldCollector
'          Call AmbilDtYgDiFU_PerFIELD
'            RPT.Reset
''            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
''            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
''            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
''            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
''            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\RptTrackingField.rpt"
'            Call SHOW_PRN
'
'     Case 13
'            WaitSecs (2)
'            Call Isi_Agent_mgm
'            Call hitung_BatchCallInitilized_PerAgent_mgm
'            Call hitung_BatchCallInitilized_PerAgent_Compare
'            Call AmbilDtYgDiFU_PerAgent
'            Call AmbilDtYgDiFU_PerAgent_Compare
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ChartUtilizedCallAgent.rpt"
'            Call SHOW_PRN
'
'     Case 14
'            WaitSecs (2)
'            Call Isi_Agent_mgm
'            Call AmbilDtYgDiFU_PerAgent
'            Call AmbilDtYgDiFU_PerAgent_Compare
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
'            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ChartUtilized.rpt"
'            Call SHOW_PRN
'
'     Case 15
'            WaitSecs (2)
'            Call Isi_Agent_mgm
'            Call Hitung_Volume_of_Payment
'            Call Hitung_Volume_of_Payment_Compare
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
'            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPayment.rpt"
'            Call SHOW_PRN
'
'     Case 16
'            WaitSecs (2)
'            Call Isi_Agent_mgm
'            Call ReportPTPNego
'            Call ReportPTPNego_Compare
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
'            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPTP.rpt"
'            Call SHOW_PRN
'    Case 17
'            WaitSecs (2)
'            Call Isi_Agent_mgm
'            Call ReportPTPNego
'            Call ReportPTPNego_Compare
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
'            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPTP.rpt"
'            Call SHOW_PRN
'    Case 18
'            WaitSecs (2)
'            Call Isi_Agent_mgm
'            Call ReportPTPNego
'            Call ReportPTPNego_Compare
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
'            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPTP.rpt"
'            Call SHOW_PRN
'    Case 20
'            'Call Isi_AgentPTP
'            Call isi_PTP
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Detail PTP.rpt"
'            Call SHOW_PRN
'
            'Report Telepon Agent buat OT
            Case 18
                Call report_ot_agent
                Call report_ot_isi_previous
                WaitSecs (2)
                RPT.Reset
                RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
                RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
                RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
                RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\rptot.rpt"
                Call SHOW_PRN
           Case 19
           Dim rstrup As New ADODB.Recordset
           Dim rsdes As New ADODB.Recordset
           WaitSecs (2)
           RPT.Reset
           If Option1(0).Value = True Then
                If Len(SYARAT) > 0 Then
                    SYARAT = SYARAT + " and USERID between  '" + Combo2(0).text + "' and '" + Combo2(1).text + "' "
                Else
                    SYARAT = " WHERE USERID between  '" + Combo2(0).text + "' and '" + Combo2(1).text + "' "
                End If
           
           Else
           
          
           If Len(SYARAT) > 0 Then
                    SYARAT = SYARAT + " and USERID in (select userid from usertbl where "
                    SYARAT = SYARAT + " SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')"

                Else
                    SYARAT = SYARAT + " where  USERID in (select userid from usertbl where "
                    SYARAT = SYARAT + " SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')"
                End If
           End If
           
            
            
                If Not (TDBDate1(0).ValueIsNull) And Not (TDBDate1(1).ValueIsNull) Then
                    If Len(SYARAT) > 0 Then
                        SYARAT = SYARAT + " AND startdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
                    Else
                        SYARAT = " WHERE  startdate  between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
                    End If
                End If
                
                If Combo1(0).text <> "" And Combo1(1).text <> "" Then
                    If Len(SYARAT) > 0 Then
                        SYARAT = SYARAT + " AND recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
                    Else
                         SYARAT = SYARAT + " where  recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
                    End If
                End If
                
                Strsql = "select custid,count(custid) AS JML,userid as agent,date(startdate),recsource ,namach from tblphonemonitorhst" + SYARAT + " group by custid,USERID ,recsource,date(startdate),namach order by USERID,date(startdate) "
                Set rstrup = New ADODB.Recordset
                rstrup.CursorLocation = adUseClient
                rstrup.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
                
                M_RPTCONN.execute "DELETE FROM tblreportcall  "
                Strsql = "select * from tblreportcall"
                
                Set rsdes = New ADODB.Recordset
                rsdes.CursorLocation = adUseClient
                rsdes.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
                  ProgressBar1.Max = rstrup.RecordCount + 1
                While Not rstrup.EOF
               ProgressBar1.Value = rstrup.Bookmark
                      rsdes.AddNew
                      rsdes!agent = IIf(IsNull(rstrup!agent), "", rstrup!agent)
                      rsdes!CustId = IIf(IsNull(rstrup!CustId), "", rstrup!CustId)
                      rsdes!Date = IIf(IsNull(rstrup!Date), Null, rstrup!Date)
                      rsdes!jml = IIf(IsNull(rstrup!jml), 0, rstrup!jml)
                      rsdes!chname = IIf(IsNull(rstrup!namach), "", rstrup!namach)
                      rsdes!RECSOURCE = IIf(IsNull(rstrup!RECSOURCE), "", rstrup!RECSOURCE)
                      rsdes.update
                      rstrup.MoveNext
                Wend
                
                
                RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
                  RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
                RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\rptcall.rpt"
                Call SHOW_PRN
                
        '@@ 17-03-2011 Report Contacto
        Case 20
            Call IsiAgentContactto
            Call IsiContactto
            Call IsiContacttoJmlAcc
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptContactto.rpt"
            MDIForm1.TimerTanda.Enabled = False
            MDIForm1.ShapeTanda.FillColor = vbBlack
            Call SHOW_PRN
            
        '@@ 15-03-2012 Report Hot Prospect
        Case 27
            Call Isi_Hot_Prospect
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptHotProspect.rpt"
            Call SHOW_PRN
            
        '@@30-03-2012 Report Keep Account
        Case 28
            Call Isi_Keep_Account
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).text & " " & DTimeLastCall(0).text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).text & " " & DTimeLastCall(1).text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptKeepAcc.rpt"
            Call SHOW_PRN
    End Select
    ProgressBar1.Visible = False
    Case 1
        Unload Me
End Select
ProgressBar1.Visible = False


Exit Sub
Command1_ClickeR:
    MsgBox err.Description
    Resume
End Sub

Private Sub Isi_AgentPTP()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from TrackingNEGOPTP"

M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
m_objrs1.Open "Select * from TrackingNEGOPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    m_objrs1.AddNew
    m_objrs1!TEAM = CStr(IIf(IsNull(M_Objrs!TEAM), "", M_Objrs!TEAM))
    m_objrs1!TSRNAME = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!TEAM = CStr(IIf(IsNull(M_Objrs!SPVCODE), "", M_Objrs!SPVCODE))
    m_objrs1!aoc = CStr(IIf(IsNull(M_Objrs!Userid), "", M_Objrs!Userid))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
End Sub

Private Sub isi_PTP()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim JUMLAH As String
Dim agent, CustId, TEAM, Name As String
Dim tipe As String
Dim TGL As Integer
Dim cmdsql As String
Dim CMDSQL1 As String
Dim TGLJanji As Date
Dim x%
Dim Jns As String

On Error GoTo hitung_JmlDataer

M_RPTCONN.execute "DELETE FROM TrackingNEGOPTP"

M_Objrs.CursorLocation = adUseClient

If Check1.Value = 1 Then
    Jns = "REG"
ElseIf Check2.Value = 1 Then
    Jns = "IPO"
ElseIf Check3.Value = 1 Then
    Jns = "RPO"
End If
'CMDSQL = "SELECT mgm.AGENT,usertbl.TEAM,mgm.CUSTID,mgm.NAME,TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.Promisepay from usertbl,TBLNEGOPTP,mgm "
'CMDSQL = CMDSQL + "Where mgm.CustId = TBLNEGOPTP.CustId AND usertbl.USERID=mgm.AGENT AND tblnegoptp.promisedate between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'"

If Option1(1).Value = True Then
    If Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Then
        cmdsql = "Select * from reportPTP where TYPE='" + Jns + "' and promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') ORDER BY NAME"
    Else
        cmdsql = "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') ORDER BY NAME"
    End If
    'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        If Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Then
            cmdsql = "Select * from reportPTP where TYPE='" + Jns + "' and promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' ORDER BY NAME"
        Else
            cmdsql = "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' ORDER BY NAME"
        'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        End If
    End If
End If

M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    TEAM = CStr(IIf(IsNull(M_Objrs!TEAM), "", M_Objrs!TEAM))
    tipe = CStr(IIf(IsNull(M_Objrs!Type), "", M_Objrs!Type))
    TGLJanji = CStr(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate))
    cmdsql = "Insert into TrackingNEGOPTP (AOC,Custid,NAME,TEAM,TYPE,PROMISEDATE) values "
    cmdsql = cmdsql + "('" & agent & "','" & CustId & "','" & Name & "','" & TEAM & "','" & tipe & "','" & TGLJanji & "')"
    M_RPTCONN.execute cmdsql
    M_Objrs.MoveNext
Wend

If M_Objrs.BOF And M_Objrs.EOF Then
Else
    M_Objrs.MoveFirst
    While Not M_Objrs.EOF
        agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
        CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
        TGL = IIf(IsNull(M_Objrs!PromiseDate), "", DatePart("D", M_Objrs!PromiseDate))
        TGLJanji = Format(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate), "mm/dd/yyyy")
        JUMLAH = IIf(IsNull(M_Objrs!PromisePay), "", M_Objrs!PromisePay)
        cmdsql = "update TrackingNEGOPTP set TGL" & TGL & "='" & JUMLAH & "' WHERE "
        'CMDSQL = CMDSQL + "AOC='" & agent & "' AND CUSTID='" & CustId & "'"
        cmdsql = cmdsql + "PROMISEDATE=#" & TGLJanji & "# AND CUSTID='" & CustId & "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
End If
Set M_Objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox err.Description
End Sub

Private Sub Isi_Settled_Payment()
Dim M_Objrs As New ADODB.Recordset
Dim LRECSOURCE As String
Dim m_msgbox As Variant


On Error GoTo hitung_JmlDataer
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient

cmdsql = "select agent, count(tgllunas) as jml from mgm where tgllunas BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_Objrs!jml), "0", M_Objrs!jml))
        LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch set SETTLED_PAYMENT =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    cmdsql = Empty
    LAgent = Empty
    JUMLAH = Empty
    
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If

End Sub
Private Sub Isi_Progess_OF_PAyment()
Dim M_Objrs As New ADODB.Recordset
Dim LRECSOURCE As String
Dim m_msgbox As Variant


On Error GoTo hitung_JmlDataer
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient

cmdsql = "select agent, count(tglPOP) as jml from mgm where tglPOP BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_Objrs!jml), "0", M_Objrs!jml))
        LAgent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
        cmdsql = "Update TrackingRptPerPrgBatch set POP2 =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    JUMLAH = Empty
    LAgent = Empty
    cmdsql = Empty
    
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If

End Sub


Private Sub Hitung_Number_of_Payment()
Dim M_Objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_Objrs.CursorLocation = adUseClient

cmdsql = "select agent, count(custid) as jml from (select distinct custid,agent from HtgNumberOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and recsource between'" + Combo1(0).text + "' and '" + Combo1(1).text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"

   ' CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where F_CEK ='PTP' AND recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND TGLINCOMING BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_Objrs!jml), "0", M_Objrs!jml))
        LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch set NPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If

End Sub

Private Sub Hitung_Volume_of_Payment()
Dim M_Objrs As New ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_Objrs.CursorLocation = adUseClient

cmdsql = "select agent, sum(Payment) as jml from (select * from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_Objrs!jml), "0", M_Objrs!jml))
        LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch set VolPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    JUMLAH = Empty
    LAgent = Empty
    cmdsql = Empty
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If
End Sub
Private Sub Hitung_Volume_of_Payment_Compare()
Dim M_Objrs As New ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_Objrs.CursorLocation = adUseClient

cmdsql = "select agent, sum(Payment) as jml from (select * from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' AND '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_Objrs!jml), "0", M_Objrs!jml))
        LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch set VolPayment_LM =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    JUMLAH = Empty
    LAgent = Empty
    cmdsql = Empty
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If
End Sub

Private Sub Hitung_Vol_PTP()
Dim M_Objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
Dim LRECSOURCE As String
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer

    M_Objrs.CursorLocation = adUseClient


If Option1(0).Value Then
cmdsql = "SELECT mgm_hst.agent, mgm_hst.F_CEK,count(mgm_hst.f_cek) as JML,  sum(mgm.ttlptp) as jmlPTP from mgm_hst INNER JOIN (SELECT custid, max(tgl) as tglmax FROM mgm_hst "
'CMDSQL = CMDSQL + "where tgl BETWEEN '1990-01-01 12:00:00 AM' and '2020-12-31 11:59:00 PM'  and "
cmdsql = cmdsql + "where date(tgl) BETWEEN '" & Format(TDBDate1(0).Value, "yyyy-mm-dd") & "' and '" & Format(TDBDate1(1).Value, "yyyy-mm-dd") & "'  and "
cmdsql = cmdsql + " custid in(select custid from mgm where agent between  '" + Combo2(0).text + "' and  '" + Combo2(1).text + "') and "
cmdsql = cmdsql + " f_cek_new='PTP' group by custid) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax inner join mgm on mgm.custid = a.custid and "
cmdsql = cmdsql + "mgm_hst.agent=mgm.agent where recsource BETWEEN'" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
cmdsql = cmdsql + "group by mgm_hst.agent, mgm_hst.f_cek "
cmdsql = cmdsql + "order by mgm_hst.agent"
Else
If Option1(1).Value Then
cmdsql = "SELECT mgm_hst.agent, mgm_hst.f_cek,count(mgm_hst.f_cek) as JML, sum(mgm.ttlptp) as jmlPTP from mgm_hst INNER JOIN (SELECT custid, max(tgl) as tglmax FROM mgm_hst "
'CMDSQL = CMDSQL + "where tgl BETWEEN ' 1990-01-01 12:00:00 AM' and '2020-12-31 11:59:00 PM'  and "
cmdsql = cmdsql + "where date(tgl) BETWEEN '" & Format(TDBDate1(0).Value, "yyyy-mm-dd") & "' and '" & Format(TDBDate1(1).Value, "yyyy-mm-dd") & "'  and "
cmdsql = cmdsql + " custid in(select custid from mgm where agent in (select userid from usertbl where SPVCODE >= '" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')) and "
cmdsql = cmdsql + " f_cek_new='PTP' group by custid) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax inner join mgm on mgm.custid = a.custid and "
cmdsql = cmdsql + "mgm_hst.agent=mgm.agent where recsource BETWEEN'" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
cmdsql = cmdsql + "group by mgm_hst.agent, mgm_hst.f_cek "
cmdsql = cmdsql + "order by mgm_hst.agent"
End If
End If
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_Objrs!jmlPTP), "0", M_Objrs!jmlPTP))
        LAgent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
        cmdsql = "Update TrackingRptPerPrgBatch set VolPTP ='" + JUMLAH + "' where AOC ='" + LAgent + "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If

End Sub

Private Sub Isi_Report_PTP_Jatuh_Tempo()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Isi_AgentErr


M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    cmdsql = "select * from ("
    cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') ORDER BY AGENT"
    cmdsql = cmdsql + ") as a,mgm, "
    
    '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
    cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
    cmdsql = cmdsql + " and  '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
    cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
    cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
    cmdsql = cmdsql + " SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
    cmdsql = cmdsql + " group by custid "
    cmdsql = cmdsql + ") as b "
    
    cmdsql = cmdsql + " where a.custid=mgm.custid "
    cmdsql = cmdsql + " and a.promisedate is not null "
    
    cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
    cmdsql = cmdsql + " and b.custid=mgm.custid "
    '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
    'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
    'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
    

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        cmdsql = "select * from ("
        cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' ORDER BY AGENT"
        cmdsql = cmdsql + ") as a,mgm, "
        
        '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
        cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
        cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
        cmdsql = cmdsql + " and  '"
        cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
        cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
        cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
        cmdsql = cmdsql + " userid between '" + Combo2(0).text + "' and  '" + Combo2(1).text + "') "
        cmdsql = cmdsql + " group by custid "
        cmdsql = cmdsql + ") as b "
        
        cmdsql = cmdsql + " where a.custid=mgm.custid "
        cmdsql = cmdsql + " and a.promisedate is not null "
        
        cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
        cmdsql = cmdsql + " and b.custid=mgm.custid "
        '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
        'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
        'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate))
    If IsNull(M_Objrs!tglallptp) Then
        m_objrs1!tglallptp = Null
    Else
        m_objrs1!tglallptp = CStr(IIf(IsNull(M_Objrs!tglallptp), "", M_Objrs!tglallptp))
    End If
    
    m_objrs1!inputdate = CStr(IIf(IsNull(M_Objrs!inputdate), "2020-12-30", M_Objrs!inputdate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!Principal), "0", M_Objrs!Principal))
    m_objrs1!ptpvia = IIf(IsNull(M_Objrs!ptpvia), "", M_Objrs!ptpvia)
    
    '@@26-01-2012 Tambahan Tanggal Tagih dan result
    If IsNull(M_Objrs!tgl_tagih) Then
        m_objrs1!tgl_tagih = Null
    Else
        m_objrs1!tgl_tagih = CStr(IIf(IsNull(M_Objrs!tgl_tagih), "", M_Objrs!tgl_tagih))
    End If
    
    m_objrs1!result = IIf(IsNull(M_Objrs!result_ptp), "", M_Objrs!result_ptp)
    
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume

End Sub

'Randy 26March2015
Private Sub Isi_Report_PTP_REG_Jatuh_Tempo()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Isi_AgentErr


M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    cmdsql = "select * from ("
    cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') AND tenor > 1 ORDER BY AGENT"
    cmdsql = cmdsql + ") as a,mgm, "
    
    '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
    cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
    cmdsql = cmdsql + " and  '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
    cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
    cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
    cmdsql = cmdsql + " SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
    cmdsql = cmdsql + " group by custid "
    cmdsql = cmdsql + ") as b "
    
    cmdsql = cmdsql + " where a.custid=mgm.custid "
    cmdsql = cmdsql + " and a.promisedate is not null "
    
    cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
    cmdsql = cmdsql + " and b.custid=mgm.custid "
    '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
    'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
    'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
    

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        cmdsql = "select * from ("
        cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' AND tenor > 1 ORDER BY AGENT"
        cmdsql = cmdsql + ") as a,mgm, "
        
        '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
        cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
        cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
        cmdsql = cmdsql + " and  '"
        cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
        cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
        cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
        cmdsql = cmdsql + " userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "') "
        cmdsql = cmdsql + " group by custid "
        cmdsql = cmdsql + ") as b "
        
        cmdsql = cmdsql + " where a.custid=mgm.custid "
        cmdsql = cmdsql + " and a.promisedate is not null "
        
        cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
        cmdsql = cmdsql + " and b.custid=mgm.custid "
        '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
        'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
        'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!Tenor = CStr(IIf(IsNull(M_Objrs!Tenor), "", M_Objrs!Tenor))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate))
    If IsNull(M_Objrs!tglallptp) Then
        m_objrs1!tglallptp = Null
    Else
        m_objrs1!tglallptp = CStr(IIf(IsNull(M_Objrs!tglallptp), "", M_Objrs!tglallptp))
    End If
    
    m_objrs1!inputdate = CStr(IIf(IsNull(M_Objrs!inputdate), "2020-12-30", M_Objrs!inputdate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!Principal), "0", M_Objrs!Principal))
    m_objrs1!ptpvia = IIf(IsNull(M_Objrs!ptpvia), "", M_Objrs!ptpvia)
    
    '@@26-01-2012 Tambahan Tanggal Tagih dan result
    If IsNull(M_Objrs!tgl_tagih) Then
        m_objrs1!tgl_tagih = Null
    Else
        m_objrs1!tgl_tagih = CStr(IIf(IsNull(M_Objrs!tgl_tagih), "", M_Objrs!tgl_tagih))
    End If
    
    m_objrs1!result = IIf(IsNull(M_Objrs!result_ptp), "", M_Objrs!result_ptp)
    
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume
End Sub


Private Sub Isi_Report_PTP_REGULER2()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
'On Error GoTo Isi_AgentErr


M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then
    cmdsql = " select * from ("
    cmdsql = cmdsql + "SELECT * FROM tblnegoptp_reguler WHERE promisedate between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'  "
    cmdsql = cmdsql + "and '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and agent in ("
    cmdsql = cmdsql + "select userid from usertbl where  SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')) a "
    cmdsql = cmdsql + "inner join mgm b on a.custid = b.custid where recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' order by a.agent"

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
    cmdsql = " select * from ("
    cmdsql = cmdsql + "SELECT * FROM tblnegoptp_reguler WHERE promisedate between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "''  "
    cmdsql = cmdsql + "and '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and agent in ("
    cmdsql = cmdsql + "select userid from usertbl where  userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "')) a"
    cmdsql = cmdsql + "inner join mgm b on a.custid = b.custid where recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' order by a.agent"

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!balance), "0", M_Objrs!balance))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!inputdate = CStr(IIf(IsNull(M_Objrs!inputdate), "2020-12-30", M_Objrs!inputdate))
    m_objrs1!result = IIf(IsNull(M_Objrs!Type), "", M_Objrs!Type)
    m_objrs1!Tenor = CStr(IIf(IsNull(M_Objrs!Tenor), "", M_Objrs!Tenor))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!down_payment), "0", M_Objrs!down_payment))
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    
    If IsNull(M_Objrs!tglallptp) Then
        m_objrs1!tglallptp = Null
    Else
        m_objrs1!tglallptp = CStr(IIf(IsNull(M_Objrs!tglallptp), "", M_Objrs!tglallptp))
    End If
    
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

Exit Sub

'Isi_AgentErr:
    MsgBox err.Description
    'Resume

End Sub


Private Sub Isi_Report_On_Going_PTP()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
'On Error GoTo Isi_AgentErr


M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then
    cmdsql = " select * from ("
    cmdsql = cmdsql + "SELECT * FROM tblnegoptp_reguler WHERE promisedate between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'  "
    cmdsql = cmdsql + "and '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and agent in ("
    cmdsql = cmdsql + "select userid from usertbl where  SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')) a "
    cmdsql = cmdsql + "inner join mgm b on a.custid = b.custid where recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' order by a.agent"

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
    cmdsql = " select * from ("
    cmdsql = cmdsql + "SELECT * FROM tblnegoptp_reguler WHERE promisedate between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "''  "
    cmdsql = cmdsql + "and '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and agent in ("
    cmdsql = cmdsql + "select userid from usertbl where  userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "')) a"
    cmdsql = cmdsql + "inner join mgm b on a.custid = b.custid where recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' order by a.agent"

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!balance), "0", M_Objrs!balance))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!inputdate = CStr(IIf(IsNull(M_Objrs!inputdate), "2020-12-30", M_Objrs!inputdate))
    m_objrs1!result = IIf(IsNull(M_Objrs!Type), "", M_Objrs!Type)
    m_objrs1!Tenor = CStr(IIf(IsNull(M_Objrs!Tenor), "", M_Objrs!Tenor))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!down_payment), "0", M_Objrs!down_payment))
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    
    If IsNull(M_Objrs!tglallptp) Then
        m_objrs1!tglallptp = Null
    Else
        m_objrs1!tglallptp = CStr(IIf(IsNull(M_Objrs!tglallptp), "", M_Objrs!tglallptp))
    End If
    
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

Exit Sub

'Isi_AgentErr:
    MsgBox err.Description
    'Resume

End Sub

'REQUEST DODDY 18MEI2015(Randy)
Private Sub Isi_Report_Result_DeskCall()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim cmdsql As String
Dim iQuery As String
'On Error GoTo Isi_AgentErr

M_OBJCONN.execute "DROP TABLE IF EXISTS tbltemp_pergerakan_status"
'jejaktian 30022016
'If Option1(1).Value = True Then
'    iQuery = " CREATE TABLE tbltemp_pergerakan_status AS  SELECT * FROM ( "
'    iQuery = iQuery + " SELECT b.agent, a.agent as id_agent,custid, status_call_sebelum, status_call_sekarang, tglcall, team as nama_tl, waktu_mulai_call"
'    iQuery = iQuery + " FROM ("
'    iQuery = iQuery + " SELECT agent,custid, status_call_sebelum, f_cek_new as status_call_sekarang, tglcall, waktu_mulai_call"
'    iQuery = iQuery + " FROM mgm WHERE tglcall between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'"
'    iQuery = iQuery + " AND '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
'    iQuery = iQuery + " AND agent in("
'    iQuery = iQuery + " SELECT userid FROM usertbl WHERE spvcode >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' order by spvcode)) as a "
'    iQuery = iQuery + " left join usertbl b on a.agent = b.userid) as c left join (SELECT userid, team from usertbl) d on c.id_agent = d.userid"
'Else
'    iQuery = " CREATE TABLE tbltemp_pergerakan_status AS  SELECT * FROM ( "
'    iQuery = iQuery + " SELECT b.agent, a.agent as id_agent,custid, status_call_sebelum, status_call_sekarang, tglcall, team as nama_tl, waktu_mulai_call"
'    iQuery = iQuery + " FROM ("
'    iQuery = iQuery + " SELECT agent,custid, status_call_sebelum, f_cek_new as status_call_sekarang, tglcall, waktu_mulai_call"
'    iQuery = iQuery + " FROM mgm WHERE tglcall between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'"
'    iQuery = iQuery + " AND '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
'    iQuery = iQuery + " AND agent in("
'    iQuery = iQuery + " SELECT userid FROM usertbl WHERE userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' order by spvcode)) as a "
'    iQuery = iQuery + " left join usertbl b on a.agent = b.userid) as c left join (SELECT userid, team from usertbl) d on c.id_agent = d.userid"
'End If
'M_OBJCONN.Execute iQuery
'========================================================
'jejaktian08032016

If Option1(1).Value = True Then
    iQuery = " CREATE TABLE tbltemp_pergerakan_status AS  SELECT * FROM ( "
    iQuery = iQuery + " SELECT b.agent, a.agent as id_agent,custid, sstatus_awal, sstatus_akhir, team as nama_tl, start_time, stop_time"
    iQuery = iQuery + " FROM ("
    iQuery = iQuery + " SELECT agent,custid, sstatus_awal, sstatus_akhir,start_time,stop_time"
    iQuery = iQuery + " FROM tblrrd WHERE start_time between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'"
    iQuery = iQuery + " AND '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
    iQuery = iQuery + " AND agent in("
    iQuery = iQuery + " SELECT userid FROM usertbl WHERE spvcode >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "' order by spvcode)) as a "
    iQuery = iQuery + " left join usertbl b on a.agent = b.userid) as c left join (SELECT userid, team from usertbl) d on c.id_agent = d.userid"
Else
    iQuery = " CREATE TABLE tbltemp_pergerakan_status AS  SELECT * FROM ( "
    iQuery = iQuery + " SELECT b.agent, a.agent as id_agent,custid, sstatus_awal, sstatus_akhir, team as nama_tl, start_time, stop_time"
    iQuery = iQuery + " FROM ("
    iQuery = iQuery + " SELECT agent,custid, sstatus_awal, sstatus_akhir,start_time,stop_time"
    iQuery = iQuery + " FROM tblrrd WHERE start_time between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'"
    iQuery = iQuery + " AND '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
    iQuery = iQuery + " AND agent in("
    iQuery = iQuery + " SELECT userid FROM usertbl WHERE userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' order by spvcode)) as a "
    iQuery = iQuery + " left join usertbl b on a.agent = b.userid) as c left join (SELECT userid, team from usertbl) d on c.id_agent = d.userid"
End If
M_OBJCONN.execute iQuery

M_RPTCONN.execute "Delete * from TrackingRptPerPrgBatch"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

cmdsql = " SELECT * FROM tbltemp_pergerakan_status "
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText


m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs1.RecordCount + 1

While Not M_Objrs.EOF
    'ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!TEAM = CStr(IIf(IsNull(M_Objrs!nama_tl), "", M_Objrs!nama_tl))
    m_objrs1!TSRNAME = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!status_call_sebelum = IIf(IsNull(M_Objrs!sstatus_awal), "", M_Objrs!sstatus_awal)
    m_objrs1!status_call_sekarang = CStr(IIf(IsNull(M_Objrs!sstatus_akhir), "", M_Objrs!sstatus_akhir))
    m_objrs1!aoc = CStr(IIf(IsNull(M_Objrs!Userid), "", M_Objrs!Userid))
    m_objrs1!tgl_call = CStr(IIf(IsNull(M_Objrs!stop_time), "1900-01-01", Format(M_Objrs!stop_time, "yyyy-mm-dd hh:mm:ss")))
    m_objrs1!tgl_mulai_call = CStr(IIf(IsNull(M_Objrs!start_time), "1900-01-01", Format(M_Objrs!start_time, "yyyy-mm-dd hh:mm:ss")))

    
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

Exit Sub

'Isi_AgentErr:
    MsgBox err.Description
    'Resume
End Sub




'Randy 26March2015
Private Sub Isi_Report_PTP_REG_Jatuh_Tempo_Excel()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Isi_AgentErr


M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    cmdsql = "select * from ("
    cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') AND tenor > 1 ORDER BY AGENT"
    cmdsql = cmdsql + ") as a,mgm, "
    
    '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
    cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
    cmdsql = cmdsql + " and  '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
    cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
    cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
    cmdsql = cmdsql + " SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
    cmdsql = cmdsql + " group by custid "
    cmdsql = cmdsql + ") as b "
    
    cmdsql = cmdsql + " where a.custid=mgm.custid "
    cmdsql = cmdsql + " and a.promisedate is not null "
    
    cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
    cmdsql = cmdsql + " and b.custid=mgm.custid "
    '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
    'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
    'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
    

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        cmdsql = "select * from ("
        cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' AND tenor > 1 ORDER BY AGENT"
        cmdsql = cmdsql + ") as a,mgm, "
        
        '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
        cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
        cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
        cmdsql = cmdsql + " and  '"
        cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
        cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
        cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
        cmdsql = cmdsql + " userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "') "
        cmdsql = cmdsql + " group by custid "
        cmdsql = cmdsql + ") as b "
        
        cmdsql = cmdsql + " where a.custid=mgm.custid "
        cmdsql = cmdsql + " and a.promisedate is not null "
        
        cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
        cmdsql = cmdsql + " and b.custid=mgm.custid "
        '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
        'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
        'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!Tenor = CStr(IIf(IsNull(M_Objrs!Tenor), "", M_Objrs!Tenor))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate))
    If IsNull(M_Objrs!tglallptp) Then
        m_objrs1!tglallptp = Null
    Else
        m_objrs1!tglallptp = CStr(IIf(IsNull(M_Objrs!tglallptp), "", M_Objrs!tglallptp))
    End If
    
    m_objrs1!inputdate = CStr(IIf(IsNull(M_Objrs!inputdate), "2020-12-30", M_Objrs!inputdate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!Principal), "0", M_Objrs!Principal))
    m_objrs1!ptpvia = IIf(IsNull(M_Objrs!ptpvia), "", M_Objrs!ptpvia)
    
    '@@26-01-2012 Tambahan Tanggal Tagih dan result
    If IsNull(M_Objrs!tgl_tagih) Then
        m_objrs1!tgl_tagih = Null
    Else
        m_objrs1!tgl_tagih = CStr(IIf(IsNull(M_Objrs!tgl_tagih), "", M_Objrs!tgl_tagih))
    End If
    
    m_objrs1!result = IIf(IsNull(M_Objrs!result_ptp), "", M_Objrs!result_ptp)
    
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing


    If b_excel Then
        'Call UpdateAllPaymentCPA
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew_Excel_Reguler.rpt"
        RPT.RetrieveDataFiles
        RPT.Destination = crptToFile
        RPT.PrintFileType = crptExcel50
        RPT.action = 1
    End If
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume

End Sub

'Randy 26March2015
Private Sub Isi_Report_PTP_PO_Jatuh_Tempo()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Isi_AgentErr


M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    cmdsql = "select * from ("
    cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') AND tenor <= 1 ORDER BY AGENT"
    cmdsql = cmdsql + ") as a,mgm, "
    
    '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
    cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
    cmdsql = cmdsql + " and  '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
    cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
    cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
    cmdsql = cmdsql + " SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
    cmdsql = cmdsql + " group by custid "
    cmdsql = cmdsql + ") as b "
    
    cmdsql = cmdsql + " where a.custid=mgm.custid "
    cmdsql = cmdsql + " and a.promisedate is not null "
    
    cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
    cmdsql = cmdsql + " and b.custid=mgm.custid "
    '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
    'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
    'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
    

    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        cmdsql = "select * from ("
        cmdsql = cmdsql + "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' AND tenor <= 1 ORDER BY AGENT"
        cmdsql = cmdsql + ") as a,mgm, "
        
        '@@27-06-2012 Ambil berdasarkan tanggal terakhir negoptp
        cmdsql = cmdsql + " (select custid,max(promisedate) as tglakhir from reportPTP where promisedate between '"
        cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' "
        cmdsql = cmdsql + " and  '"
        cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
        cmdsql = cmdsql + " recsource between  '" + Combo1(0).text + "' and '"
        cmdsql = cmdsql + Combo1(1).text + "' and agent in (select userid from usertbl where "
        cmdsql = cmdsql + " userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "') "
        cmdsql = cmdsql + " group by custid "
        cmdsql = cmdsql + ") as b "
        
        cmdsql = cmdsql + " where a.custid=mgm.custid "
        cmdsql = cmdsql + " and a.promisedate is not null "
        
        cmdsql = cmdsql + " and a.custid=b.custid and a.promisedate=b.tglakhir "
        cmdsql = cmdsql + " and b.custid=mgm.custid "
        '@@ 16-April-2012 , ditambahkan filter bahwa custid yang jatuh tempo adalah
        'custid yang tidak berada di posisi lunas (custid tersebut tidak berada di vwwlunas)
        'CMDSQL = CMDSQL + " and a.custid not in (select custid from vwwlunas where custid <>'')"
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "", M_Objrs!PromiseDate))
    If IsNull(M_Objrs!tglallptp) Then
        m_objrs1!tglallptp = Null
    Else
        m_objrs1!tglallptp = CStr(IIf(IsNull(M_Objrs!tglallptp), "", M_Objrs!tglallptp))
    End If
    
    m_objrs1!inputdate = CStr(IIf(IsNull(M_Objrs!inputdate), "2020-12-30", M_Objrs!inputdate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!Principal), "0", M_Objrs!Principal))
    m_objrs1!ptpvia = IIf(IsNull(M_Objrs!ptpvia), "", M_Objrs!ptpvia)
    
    '@@26-01-2012 Tambahan Tanggal Tagih dan result
    If IsNull(M_Objrs!tgl_tagih) Then
        m_objrs1!tgl_tagih = Null
    Else
        m_objrs1!tgl_tagih = CStr(IIf(IsNull(M_Objrs!tgl_tagih), "", M_Objrs!tgl_tagih))
    End If
    
    m_objrs1!result = IIf(IsNull(M_Objrs!result_ptp), "", M_Objrs!result_ptp)
    
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume
End Sub

Private Sub Isi_Report_FormVisit()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from FormVisit"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(0).Value = True Then
cmdsql = "SELECT TblVisit.*, mgm.Principal AS PRINCIPLE, mgm.AmountWo AS AmountWO,mgm.name as NAME "
cmdsql = cmdsql + "FROM mgm INNER JOIN "
cmdsql = cmdsql + "TblVisit ON dbo.mgm.CUSTID = dbo.TblVisit.CUSTID "
cmdsql = cmdsql + "WHERE TblVisit.agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' "
cmdsql = cmdsql + "AND tblvisit.requestDate between  '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' "
cmdsql = cmdsql + " AND sts = '0'"
cmdsql = cmdsql + " ORDER BY tblvisit.VisitNo"
'CMDSQL = "Select * from mgm where f_cek='PTP' and tdbdatePTP between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') ORDER BY AGENT"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(1).Value = True Then
cmdsql = "SELECT TblVisit.*, mgm.Principal AS PRINCIPLE, mgm.AmountWo AS AmountWO, mgm.name as NAME "
cmdsql = cmdsql + "FROM mgm INNER JOIN "
cmdsql = cmdsql + "TblVisit ON dbo.mgm.CUSTID = dbo.TblVisit.CUSTID "
cmdsql = cmdsql + "WHERE TblVisit.agent in (SELECT userid from usertbl where SPVCODE  >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
cmdsql = cmdsql + "AND tblvisit.requestDate between  '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' "
cmdsql = cmdsql + " AND sts = '0'"
cmdsql = cmdsql + " ORDER BY tblvisit.VisitNo"
       M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
m_objrs1.Open "Select * from formVisit", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
    m_objrs1!FFC = Trim(CStr(IIf(IsNull(M_Objrs!FFC), "", M_Objrs!FFC)))
    m_objrs1!CustId = Trim(CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId)))
    m_objrs1!Name = Trim(CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name)))
    m_objrs1!NoVisit = Trim(CStr(IIf(IsNull(M_Objrs!VisitNo), "", M_Objrs!VisitNo)))
    m_objrs1!RequestDate = Trim(CStr(IIf(IsNull(M_Objrs!RequestDate), "2020-12-30", M_Objrs!RequestDate)))
    m_objrs1!DetailsR = Trim(CStr(IIf(IsNull(M_Objrs!DetailsR), "0", M_Objrs!DetailsR)))
    m_objrs1!F_CEK = Trim(CStr(IIf(IsNull(M_Objrs!F_CEK), "0", M_Objrs!F_CEK)))
    m_objrs1!VisitKe = Trim(CStr(IIf(IsNull(M_Objrs!VisitKe), "0", M_Objrs!VisitKe)))
    m_objrs1!AddressToVisit = Trim(CStr(IIf(IsNull(M_Objrs!AddressToVisit), "", M_Objrs!AddressToVisit)))
    m_objrs1!Principle = Trim(CStr(IIf(IsNull(M_Objrs!Principle), "0", M_Objrs!Principle)))
    m_objrs1!amountwo = Trim(CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo)))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume

End Sub
Private Sub Isi_Report_PTP_lunas()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Isi_REportErr

M_RPTCONN.execute "Delete * from TrackingRptPayment"

M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

    cmdsql = "select mgm.dateptp,mgm.custid, mgm.name, mgm.agent, ttlPTP, jmlBayar, mgm.ttlPTP-jmlbayar as sisaPay, mgm.tdbdatePTP,usertbl.spvcode from mgm inner join(select custid, sum(payment)as jmlBayar from tbllunas  group by custid )as a on mgm.custid = a.custid INNER JOIN usertbl on usertbl.userid=mgm.agent where tglstatus between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND  spvcode between '" + Combo3(0).text + "' and '" + Combo3(1).text + "' ORDER BY mgm.agent"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs1.Open "Select * from TrackingRptPayment", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
    m_objrs1.AddNew
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!ttlptp), 0, M_Objrs!ttlptp))
    m_objrs1!jmlBayar = CStr(IIf(IsNull(M_Objrs!jmlBayar), 0, M_Objrs!jmlBayar))
    m_objrs1!SisaPay = CStr(IIf(IsNull(M_Objrs!SisaPay), 0, M_Objrs!SisaPay))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!dateptp), "2020-12-01", M_Objrs!dateptp))
    m_objrs1!SPVCODE = CStr(IIf(IsNull(M_Objrs!SPVCODE), "", M_Objrs!SPVCODE))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_REportErr:
    MsgBox err.Description
    'Resume
End Sub

Private Sub Isi_Agent_mgm()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
'On Error GoTo Isi_AgentErr
Dim cmdsql As String
M_RPTCONN.execute "Delete * from TrackingRptPerPrgBatch"

M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    cmdsql = "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "' AND USERTYPE =1 "
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' AND USERTYPE =1 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If

m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    m_objrs1.AddNew
    m_objrs1!TEAM = Trim(CStr(IIf(IsNull(M_Objrs!TEAM), "", M_Objrs!TEAM)))
    m_objrs1!TSRNAME = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
    m_objrs1!TEAM = Trim(CStr(IIf(IsNull(M_Objrs!SPVCODE), "", M_Objrs!SPVCODE)))
    m_objrs1!aoc = Trim(CStr(IIf(IsNull(M_Objrs!Userid), "", M_Objrs!Userid)))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume
End Sub
Private Sub Isi_Field_Collector()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from TrackingRptField"

M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "' AND USERTYPE =2 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' AND USERTYPE =2 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        M_Objrs.Open "Select * from usertbl where AKTIF = 0 AND USERTYPE =2", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingRptField", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    m_objrs1.AddNew
    m_objrs1!TEAM = Trim(CStr(IIf(IsNull(M_Objrs!TEAM), "", M_Objrs!TEAM)))
    m_objrs1!TSRNAME = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
    m_objrs1!TEAM = Trim(CStr(IIf(IsNull(M_Objrs!SPVCODE), "", M_Objrs!SPVCODE)))
    m_objrs1!aoc = Trim(CStr(IIf(IsNull(M_Objrs!Userid), "", M_Objrs!Userid)))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume
End Sub

Private Sub Tracking_PTP_JatuhTempo_NEW()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
Dim jenis As String
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Check1.Value = 1 Then
jenis = "REG"
ElseIf Check2.Value = 1 Then
jenis = "IPO"

ElseIf Check3.Value = 1 Then
jenis = "RPO"
Else
'jenis = "REG"
End If

If Option1(1).Value = True Then
cmdsql = "Select * from reportptpnew where f_cek like '%PTP%' and promisedate  "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
cmdsql = cmdsql + " ORDER BY AGENT"
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
cmdsql = "Select * from reportptpnew where f_cek like '%PTP%' and promisedate  "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "' "
cmdsql = cmdsql + " ORDER BY AGENT "
'
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!F_CEK = CStr(IIf(IsNull(M_Objrs!F_CEK), "0", M_Objrs!F_CEK))
    
   ' M_OBJRS1!BaseOn = CStr(IIf(IsNull(M_OBJRS!CmbBaseOn), "", M_OBJRS!CmbBaseOn))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!Principal), "0", M_Objrs!Principal))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

' Update sumreserved
'cmdsql = "select custid, sum(promisepay) from reportPTP where "
'cmdsql = cmdsql + " f_cek='PTP' and promisedate   Between '2009/06/01 12:00:00 AM' and  '2009/07/31 11:59:00 PM'  and  RECSOURCE Between  '-----' and 'ZZZZZ' and"
'cmdsql = cmdsql + " agent >= '-----' and agent <= 'ZZZZZ'"
'cmdsql = cmdsql + "group by custid"

If Option1(1).Value = True Then
cmdsql = "select custid, sum(promisepay)as reservedPTP from reportPTP where f_cek like '%PTP%' and promisedate  "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') GROUP BY CUSTID"
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
cmdsql = "select custid, sum(promisepay)as reservedPTP from reportPTP where f_cek like '%PTP%' and promisedate  "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "' GROUP BY CUSTID "
'
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
Dim Lcustid As String
Dim sumreserved As String
While Not M_Objrs.EOF
Lcustid = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
sumreserved = IIf(IsNull(M_Objrs!ReservedPTP), "0", M_Objrs!ReservedPTP)
M_RPTCONN.execute "UPDATE TrackingPTP set sumreserved=" + sumreserved + " where custid='" + Trim(Lcustid) + "'"
M_Objrs.MoveNext
Wend

Set M_Objrs = Nothing
Exit Sub



Isi_AgentErr:
    MsgBox err.Description
    'Resume
End Sub
Private Sub TrackingReservedPTP()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
'Dim jenis As String
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then
cmdsql = "select custid,sum(promisepay) as ReservedPTP,count(custid) as tenor1, recsource,agent,promisedate, name,amountwo,ttlptp,tenor  from reportreserve  "
cmdsql = cmdsql + " where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
cmdsql = cmdsql + " group by custid, recsource,agent,name, promisedate,AMOUNTWO,ttlptp,TENOR"

'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
'cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') "
'cmdsql = cmdsql + " ORDER BY AGENT"
 M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
    
        cmdsql = "select custid,sum(promisepay) as  ReservedPTP, recsource,agent,promisedate, name, AMOUNTWO,ttlptp,TENOR  from reportreserve  "
        cmdsql = cmdsql + " where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' AND AGENT >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "'"
        cmdsql = cmdsql + " group by custid, recsource,agent,name, promisedate,AMOUNTWO, ttlptp,TENOR"
    
'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' "
'cmdsql = cmdsql + " ORDER BY AGENT "
'
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!sumreserved = CStr(IIf(IsNull(M_Objrs!ReservedPTP), "0", M_Objrs!ReservedPTP))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!ttlptp), "0", M_Objrs!ttlptp))
    m_objrs1!Tenor = CStr(IIf(IsNull(M_Objrs!Tenor), "0", M_Objrs!Tenor))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub
Isi_AgentErr:
MsgBox err.Description
End Sub

Private Sub GET_PTP_NEW()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
Dim jenis As String
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Check1.Value = 1 Then
jenis = "REG"
ElseIf Check2.Value = 1 Then
jenis = "IPO"

ElseIf Check3.Value = 1 Then
jenis = "RPO"
Else
'jenis = "REG"
End If

If Option1(1).Value = True Then
cmdsql = "Select * from reportptp where f_cek like '%PTP%' and INputdate "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
cmdsql = cmdsql + " ORDER BY AGENT"
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
cmdsql = "Select * from reportptp where f_cek like '%PTP%' and INputdate "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "' "
cmdsql = cmdsql + " ORDER BY AGENT "
'
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!F_CEK = CStr(IIf(IsNull(M_Objrs!F_CEK), "0", M_Objrs!F_CEK))
    
    'm_objrs1!BaseOn = CStr(IIf(IsNull(m_objrs!CmbBaseOn), "", m_objrs!CmbBaseOn))
    'm_objrs1!Principle = CStr(IIf(IsNull(m_objrs!Principal), "0", m_objrs!Principal))
    'm_objrs1!amountwo = CStr(IIf(IsNull(m_objrs!amountwo), "0", m_objrs!amountwo))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

' Update sumreserved
cmdsql = "select custid, sum(promisepay) from reportPTP where "
cmdsql = cmdsql + " f_cek='PTP' and INputdate   Between '2009/06/01 12:00:00 AM' and  '2009/07/31 11:59:00 PM'  and  RECSOURCE Between  '-----' and 'ZZZZZ' and"
cmdsql = cmdsql + " agent >= '-----' and agent <= 'ZZZZZ'"
cmdsql = cmdsql + "group by custid"

If Option1(1).Value = True Then
cmdsql = "select custid, sum(promisepay)as reservedPTP from reportPTP where f_cek like '%PTP-NE%' and INputdate "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') GROUP BY CUSTID"
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
cmdsql = "select custid, sum(promisepay)as reservedPTP from reportPTP where f_cek like '%PTP-NE%' and INputdate "
cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "' GROUP BY CUSTID "
'
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
Dim Lcustid As String
Dim sumreserved As String
While Not M_Objrs.EOF
Lcustid = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
sumreserved = IIf(IsNull(M_Objrs!ReservedPTP), "0", M_Objrs!ReservedPTP)
M_RPTCONN.execute "UPDATE TrackingPTP set sumreserved=" + sumreserved + " where custid='" + Trim(Lcustid) + "'"
M_Objrs.MoveNext
Wend

Set M_Objrs = Nothing
Exit Sub



Isi_AgentErr:
    MsgBox err.Description
    'Resume


End Sub

Private Sub hitung_JmlData_PerAgent_mgm()
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer
M_Objrs.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
cmdsql = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).text + "' and '" + Combo1(1).text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_Objrs!jml), "0", M_Objrs!jml)))
    JUMLAHVOL = Trim(CStr(IIf(IsNull(M_Objrs!JMLVOL), "0", M_Objrs!JMLVOL)))
    LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
    cmdsql = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + ", JMLVOL= '" + JUMLAHVOL + "'  where AOC ='" + LAgent + "';"
    M_RPTCONN.execute cmdsql
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
JUMLAH = Empty
JUMLAHVOL = Empty
Exit Sub
hitung_JmlDataer:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next
End Sub


Private Sub hitung_JmlData_FieldCollector()
Dim M_Objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim JUMLAHVOL As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_Objrs.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
cmdsql = "select FFC,count(FFC) as jml, sum(mgm.Amountwo) as JMLVOL from tblvisit INNER JOIN "
cmdsql = cmdsql + " mgm on tblVisit.custid=mgm.custid group by FFC "
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_Objrs!jml), "0", M_Objrs!jml)))
    JUMLAHVOL = Trim(CStr(IIf(IsNull(M_Objrs!JMLVOL), "0", M_Objrs!JMLVOL)))
    LAgent = Trim(CStr(IIf(IsNull(M_Objrs!FFC), "", M_Objrs!FFC)))
    cmdsql = "Update TrackingRptField set DATASIZE =" + JUMLAH + ", JMLVOL= " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.execute cmdsql
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Exit Sub
hitung_JmlDataer:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next
End Sub
Private Sub Hitung_TrackingReportPerAgent_mgm()
Dim M_Objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
Dim LAgent As String

On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    cmdsql = "Select AGENT, kethslkerja, count(AGENT) as jumlah from mgm where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'group by AGENT, kethslkerja"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_Objrs.RecordCount + 1
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
'         WaitSecs (0.5)
        LAgent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + "[" + Trim(CStr(M_Objrs!KETHSLKERJA)) + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(M_Objrs!JUMLAH), 0, M_Objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(M_Objrs!KETHSLKERJA) Then
        Else
            If M_Objrs!KETHSLKERJA = "[]" Then
            Else
                If M_Objrs!JUMLAH = 0 Then
                Else
                   
                    M_RPTCONN.execute cmdsql
                End If
            End If
        End If
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox err.Description
'Resume
End Sub

Private Sub hitung_BatchCallInitilized_PerAgent_mgm()
Dim M_Objrs As New ADODB.Recordset
'Dim JUMLAH As String
'Dim batch As String
'Dim CMDSQL As String
'Dim LAgent As String
Dim m_msgbox As Variant

On Error GoTo hitung_BatchCallInitilizeder
M_Objrs.CursorLocation = adUseClient
cmdsql = "Select agent,count(agent) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and custid in(select custid from mgm where recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "')group by agent order by  agent"
'CMDSQL = "Select userid,count(userid) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and left(RecSource,3) <> 'PRE' and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by userid"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_Objrs!jml), "", M_Objrs!jml))
    LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
    cmdsql = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.execute cmdsql
    End If
    M_Objrs.MoveNext
Wend
    LAgent = Empty
    cmdsql = Empty
    JUMLAH = Empty
    Set M_Objrs = Nothing
    Exit Sub
hitung_BatchCallInitilizeder:
m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
If m_msgbox = vbRetry Then
    WaitSecs (3)
    Resume
End If
End Sub
Private Sub hitung_BatchCallInitilized_PerAgent_Compare()
Dim M_Objrs As New ADODB.Recordset
Dim m_msgbox As Variant

On Error GoTo hitung_BatchCallInitilizeder
M_Objrs.CursorLocation = adUseClient
cmdsql = "Select agent,count(agent) as jml from mgm_hst where tgl between '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "'  and custid in(select custid from mgm where recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "')group by agent order by  agent"
'CMDSQL = "Select userid,count(userid) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and left(RecSource,3) <> 'PRE' and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by userid"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 2
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_Objrs!jml), "", M_Objrs!jml))
    LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
    cmdsql = "Update TrackingRptPerPrgBatch set [CALLSINITIATED_LM] ='" + JUMLAH + "' where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.execute cmdsql
    End If
    M_Objrs.MoveNext
   
Wend
Set M_Objrs = Nothing
LAgent = Empty
cmdsql = Empty
JUMLAH = Empty

Exit Sub
hitung_BatchCallInitilizeder:
m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
If m_msgbox = vbRetry Then
    WaitSecs (3)
    Resume
End If
End Sub

Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    'RPT.Action = 1
    'RPT.Reset
End Sub


Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "No", 4 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Report", 50 * TXT
End Sub

Private Sub Command2_Click()
    Form_talktime_all.Show
    Me.Hide
End Sub

Private Sub Form_Load()
Dim listItem As listItem
Dim M_Objrs As ADODB.Recordset
Set M_Objrs = New ADODB.Recordset
DTimeLastCall(0).Value = "00:00"
DTimeLastCall(1).Value = "23:59"
DTimeLastCall(2).Value = "00:00"
DTimeLastCall(3).Value = "23:59"
M_Objrs.CursorLocation = adUseClient
CmbCek.AddItem "Not Check"
CmbCek.AddItem "Accept"
CmbCek.AddItem "RETURN"
M_Objrs.Open "SELECT * FROM usertbl WHERE AKTIF = 0 ORDER BY USERID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_Objrs.EOF
    Combo2(0).AddItem M_Objrs!Userid
    Combo2(1).AddItem M_Objrs!Userid
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open "select * from datasourcetbl where substring(kodeds,1,3) <> 'PRE' order by kodeds", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_Objrs.EOF
    Combo1(0).AddItem M_Objrs!KODEDS
    Combo1(1).AddItem M_Objrs!KODEDS
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
'M_Objrs.Open "select distinct spvcode from spvtbl where type = 1 order by spvcode", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
M_Objrs.Open "select distinct SPVTBL.SPVCODE from SPVTBL, usertbl where SPVTBL.SPVCODE = usertbl.SPVCODE AND USERTYPE = '6' OR USERTYPE = '11' AND TYPE = '1' ORDER BY SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 
While Not M_Objrs.EOF
    Combo3(0).AddItem M_Objrs!SPVCODE
    Combo3(1).AddItem M_Objrs!SPVCODE
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Call header
' report baru
'Set listitem = ListView1.ListItems.ADD(, , "1")
'    listitem.SubItems(1) = "CH Data Tracking Summary PerSPV Type A"
'Set listitem = ListView1.ListItems.ADD(, , "2")
'    listitem.SubItems(1) = "CH Data Tracking Summary PerDCR Name Type A"
Set listItem = ListView1.ListItems.ADD(, , "1")
    listItem.SubItems(1) = "CH Data Tracking PerTeam Name Type B"
Set listItem = ListView1.ListItems.ADD(, , "2")
    listItem.SubItems(1) = "CH Data Tracking PerDCR Name Type B"
'Set listitem = ListView1.ListItems.ADD(, , "20")
'    listitem.SubItems(1) = "Tracking Detail PTP"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
'Set listitem = ListView1.ListItems.ADD(, , "19")
'    listitem.SubItems(1) = "Status Call Data Tracking PerDCR Name Type B"
Set listItem = ListView1.ListItems.ADD(, , "3")
    listItem.SubItems(1) = "Report Distribusi"
'Set listitem = ListView1.ListItems.ADD(, , "6")
'    listitem.SubItems(1) = "Report History Call"
'Set listitem = ListView1.ListItems.ADD(, , "7")
'    listitem.SubItems(1) = "Report History Call Group By CustID"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
'Set listitem = ListView1.ListItems.ADD(, , "8")
'    listitem.SubItems(1) = "Promise To Pay Report"
'report  4 gak ke pake suruh ilangin sama mba ulan
Set listItem = ListView1.ListItems.ADD(, , "4")
    listItem.SubItems(1) = "Report Actual Pay"
 ' still o
Set listItem = ListView1.ListItems.ADD(, , "5")
    listItem.SubItems(1) = "Report PTP Jatuh Tempo"
'Randy Update
Set listItem = ListView1.ListItems.ADD(, , "41")
    listItem.SubItems(1) = "Report PTP-REG Jatuh Tempo"
Set listItem = ListView1.ListItems.ADD(, , "42")
    listItem.SubItems(1) = "Report PTP-PO Jatuh Tempo"
Set listItem = ListView1.ListItems.ADD(, , "43")
    listItem.SubItems(1) = "Report PTP-REGULER 2"
Set listItem = ListView1.ListItems.ADD(, , "44")
    listItem.SubItems(1) = "Report On-Going PTP-REG"
'REQUEST DODDY 18Mei2015(Randy)
Set listItem = ListView1.ListItems.ADD(, , "45")
    listItem.SubItems(1) = "Report Result DeskCall"
'Randy
Set listItem = ListView1.ListItems.ADD(, , "6")
    listItem.SubItems(1) = "Report PTP NEW"
'Set listitem = ListView1.ListItems.ADD(, , "11")
'    listitem.SubItems(1) = "Report Form Visit"
Set listItem = ListView1.ListItems.ADD(, , "7")
    listItem.SubItems(1) = "Report Tracking Reserved PTP"

'Set listitem = listview1.ListItems.ADD(, , "8")
 '   listitem.SubItems(1) = "Report POP BP1"
'Set listitem = listview1.ListItems.ADD(, , "9")
 '   listitem.SubItems(1) = "Report POP BP2"
'Set listitem = listview1.ListItems.ADD(, , "10")
 '   listitem.SubItems(1) = "Report POP BP3"
Set listItem = ListView1.ListItems.ADD(, , "12")
    listItem.SubItems(1) = "Report BP Month"
Set listItem = ListView1.ListItems.ADD(, , "13")
    listItem.SubItems(1) = "Report BP day"
'Set listitem = ListView1.ListItems.ADD(, , "13")
 '   listitem.SubItems(1) = "Report BP3"
Set listItem = ListView1.ListItems.ADD(, , "14")
    listItem.SubItems(1) = "Report Log PTP"
Set listItem = ListView1.ListItems.ADD(, , "15")
    listItem.SubItems(1) = "Report Log CPA list"
   
'Set listItem = ListView1.ListItems.ADD(, , "16")
'    listItem.SubItems(1) = "Report Tracking Reserved PTP versi 2"

'@@ 21 September 2011, Report nomor 17 ditiadakan
'Set listitem = ListView1.ListItems.ADD(, , "17")
'    listitem.SubItems(1) = "Report Log CPA list detail"
    
Set listItem = ListView1.ListItems.ADD(, , "18")
    listItem.SubItems(1) = "Report Over Time Agent"

Set listItem = ListView1.ListItems.ADD(, , "19")
    listItem.SubItems(1) = "Report Call activity for agent"

'Set listitem = listview1.ListItems.ADD(, , "")
 '   listitem.SubItems(1) = "------------------------------------------"
'Set listitem = ListView1.ListItems.ADD(, , "13")
'    listitem.SubItems(1) = "Performance Appraisal Call & Utilized perAgent "
'Set listitem = ListView1.ListItems.ADD(, , "14")
'    listitem.SubItems(1) = "Performance Appraisal Call & Utilized perSPV"
'Set listitem = ListView1.ListItems.ADD(, , "15")
'    listitem.SubItems(1) = "Performance Appraisal Call & Utilized All Team"
'Set listitem = ListView1.ListItems.ADD(, , "16")
'    listitem.SubItems(1) = "Report Performance Appraisal Payment & PTP PerAgent"
'Set listitem = ListView1.ListItems.ADD(, , "17")
'    listitem.SubItems(1) = "Report Performance Appraisal Payment & PTP PerSPV"
'Set listitem = ListView1.ListItems.ADD(, , "18")
'    listitem.SubItems(1) = "Report Performance Appraisal Payment & PTP All Team"

Set listItem = ListView1.ListItems.ADD(, , "20")
    listItem.SubItems(1) = "Report Contact To"
    
'@@ 05-05-2011 Report Request
Set listItem = ListView1.ListItems.ADD(, , "21")
    listItem.SubItems(1) = "Report Request BS"
Set listItem = ListView1.ListItems.ADD(, , "22")
    listItem.SubItems(1) = "Report Request EC"
Set listItem = ListView1.ListItems.ADD(, , "23")
    listItem.SubItems(1) = "Report Request OST"
Set listItem = ListView1.ListItems.ADD(, , "24")
    listItem.SubItems(1) = "Report Request Problem"
Set listItem = ListView1.ListItems.ADD(, , "25")
    listItem.SubItems(1) = "Report Request PUM"
Set listItem = ListView1.ListItems.ADD(, , "26")
    listItem.SubItems(1) = "Report Request RS"
    
'@@ 15-03-2012, Buat Report Hot Prospect
Set listItem = ListView1.ListItems.ADD(, , "27")
    listItem.SubItems(1) = "Report Hot Prospect"
    
'@@ 30-03-2012, Buat Report Keep Account
Set listItem = ListView1.ListItems.ADD(, , "28")
    listItem.SubItems(1) = "Report Keep Account"
'@@ 14-05-2012, Buat Report UnValid No.Telepon
Set listItem = ListView1.ListItems.ADD(, , "29")
    listItem.SubItems(1) = "Report UnValid Number"
Set listItem = ListView1.ListItems.ADD(, , "30")
    listItem.SubItems(1) = "Report Valid Number"
Set listItem = ListView1.ListItems.ADD(, , "31")
    listItem.SubItems(1) = "Report Acc Review"
    
'@@--------- Report 32 Dinonaktifkan ---------------------------
'Set listitem = listview1.ListItems.ADD(, , "32")
'    listitem.SubItems(1) = "Report Send PTP Approve"
    
'@@--------- Report 33 Dinonaktifkan ---------------------------------------
Set listItem = ListView1.ListItems.ADD(, , "33")
    listItem.SubItems(1) = "Report Send PTP Reject"
    
'@@09072012,Tambahan nih Report Contact LPD 1 dan Contact Rate
Set listItem = ListView1.ListItems.ADD(, , "34")
    listItem.SubItems(1) = "Report Account Contact LPD 1"
    
'@@--------- Report 35 Dinonaktifkan ---------------------------------------
'Set listitem = listview1.ListItems.ADD(, , "35")
'    listitem.SubItems(1) = "Report Paid Off"
    
Set listItem = ListView1.ListItems.ADD(, , "36")
    listItem.SubItems(1) = "Report Durasi Call Contact LPD Server 4"
    
Set listItem = ListView1.ListItems.ADD(, , "37")
    listItem.SubItems(1) = "Report Durasi Call Contact LPD Server 5"
    
Set listItem = ListView1.ListItems.ADD(, , "38")
    listItem.SubItems(1) = "Tarik Seluruh CPA"
    
Set listItem = ListView1.ListItems.ADD(, , "39")
    listItem.SubItems(1) = "Report Request PTP"

Set listItem = ListView1.ListItems.ADD(, , "40")
    listItem.SubItems(1) = "Report Detail Payment Interval Permonth"

End Sub


Private Sub Form_Unload(Cancel As Integer)
M_OBJCONN.Close
Set M_OBJCONN = Nothing
M_OBJCONN.Open CMDSQLOPEN
End Sub

Private Sub ListView1_Click()
    Label2.Caption = ListView1.SelectedItem.SubItems(1)
    Select Case ListView1.SelectedItem.Index
    
    Case 1
        TglEnable_No
    Case 2
         TglEnable_No
    Case 3
          TglEnable_No
    Case 4
         TglEnable_No
    
    Case 6
         TglEnable_No
    Case 7
        TglEnable_No
    Case 8
        TglEnable_No
    
    Case 10
         TglEnable_No
    Case 11
         TglEnable_No
    Case 12
         TglEnable_No
    
    Case 14
        TglEnable_No
    Case 15
        TglEnable_No
    
    Case 17
        TglEnable_OK
    Case 18
        TglEnable_OK
    Case 19
        TglEnable_OK
    Case 20
        TglEnable_OK
    Case 21
        TglEnable_OK
    End Select
End Sub
Private Sub TglEnable_OK()
        TDBDate1(2).Enabled = True
        TDBDate1(3).Enabled = True
        DTimeLastCall(2).Enabled = True
        DTimeLastCall(3).Enabled = True
End Sub
Private Sub TglEnable_No()
        TDBDate1(2).Enabled = False
        TDBDate1(3).Enabled = False
        DTimeLastCall(2).Enabled = False
        DTimeLastCall(3).Enabled = False
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        If Option1(Index).Value = False Then
            Option1(1).Value = False
        Else
            Combo2(0).Enabled = True
            Combo2(1).Enabled = True
            Combo3(0).Enabled = False
            Combo3(1).Enabled = False
        End If
    Case 1
        If Option1(Index).Value = False Then
            Option1(0).Value = False
        Else
            Combo2(0).Enabled = False
            Combo2(1).Enabled = False
            Combo3(0).Enabled = True
            Combo3(1).Enabled = True
        End If
End Select
End Sub
Public Sub Hitung_Payment()
Dim M_Objrs As New ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_Objrs.CursorLocation = adUseClient

    cmdsql = "select agent, sum(Payment) as payment from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "' group by agent"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_Objrs!Payment), "0", M_Objrs!Payment))
        LAgent = Trim(CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent)))
        cmdsql = "Update TrackingRptPerPrgBatch set Payment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    JUMLAH = Empty
    LAgent = Empty
    cmdsql = Empty
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If
End Sub
Private Sub Isi_Data_Distribusi()
Dim rs As New ADODB.Recordset

M_RPTCONN.execute "delete from distribusi"
If Option1(1).Value = True Then
cmdsql = "select tglsource, recsource,agent, count(agent)as JML,sum(AMOUNTWO) as AMOUNT  from mgm "
cmdsql = cmdsql + " where tglsource  between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") + "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") + "' and"
cmdsql = cmdsql + " agent in "
cmdsql = cmdsql + " (select userid from usertbl where spvcode between '" + Combo3(0).text + "' and '" + Combo3(1).text + "') "
cmdsql = cmdsql + "AND RECSOURCE BETWEEN '" + Combo1(0).text + "' AND '" + Combo1(1).text + "'"
cmdsql = cmdsql + " group by tglsource,agent, RECSOURCE "
End If

If Option1(0).Value = True Then
cmdsql = " select tglsource, RECSOURCE,agent, count(agent)as JML,sum(AMOUNTWO) as AMOUNT from mgm where tglsource  between"
cmdsql = cmdsql + " '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") + "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") + "' and "
cmdsql = cmdsql + " agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "' "
cmdsql = cmdsql + "AND RECSOURCE BETWEEN '" + Combo1(0).text + "' AND '" + Combo1(1).text + "'"
cmdsql = cmdsql + " GROUP BY tglsource, agent, RECSOURCE"
End If
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not rs.EOF
cmdsql = "INSERT INTO DISTRIBUSI(tglsource,USERID,RECSOURCE,jumlah,AMOUNT)"
cmdsql = cmdsql + " VALUES("
cmdsql = cmdsql + " '" + IIf(IsNull(rs!TGLSOURCE), "1900/01/01", Format(rs!TGLSOURCE, "yyyy/mm/dd")) + "', "
cmdsql = cmdsql + " '" + IIf(IsNull(rs!agent), "", rs!agent) + "', "
cmdsql = cmdsql + " '" + IIf(IsNull(rs!RECSOURCE), "", rs!RECSOURCE) + "', "
cmdsql = cmdsql + " " + CStr(IIf(IsNull(rs!jml), "0", rs!jml)) + ","
cmdsql = cmdsql + " " + CStr(IIf(IsNull(rs!Amount), "0", rs!Amount)) + ""
cmdsql = cmdsql + ")"
M_RPTCONN.execute cmdsql
rs.MoveNext
Wend
Set rs = Nothing
End Sub

Private Sub ISI_DATA_POP_BP1()

'@@Report POP-BP1 [23-11-09]-- POSTGREE
Dim a As String
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient


'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
cmdsql = "select * from vwlunasmax"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
cmdsql = " select custid,name,Amountwo,agent from MGM where agent not in ('LUNAS','PULLOUT') and "
If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
cmdsql = cmdsql + " and custid in("
cmdsql = cmdsql + " select custid from vwlunasmax where datediff('month',tglbayar,now())=1)"

If Combo1(0).text <> "" And Combo1(1).text <> "" Then
  cmdsql = cmdsql + " and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
End If

M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount > 0 Then
ProgressBar1.Max = M_Objrs.RecordCount
End If
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    cmdsql = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    cmdsql = cmdsql + " ("
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId) + "', "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name) + "', "
    cmdsql = cmdsql + " " + CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo)) + ", "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent) + "', "
    cmdsql = cmdsql + " 'POP BP1') "
    M_RPTCONN.execute cmdsql

M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
    
    
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
        cmdsql = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        cmdsql = cmdsql + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
        cmdsql = cmdsql + " group by custid "
   
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
    End If
    While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set payment = " + CStr(M_Objrs!Payment) + "  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from usertbl where  "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " userid >= '" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " userid in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
        While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set spvcode = '" + CStr(M_Objrs!SPVCODE) + "'  where agent='" + Trim(M_Objrs!Userid) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "')"
    ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent in(select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode <= '" + Combo3(1).text + "'))"
    End If
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
   While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set PaymentDate = '" + Format(M_Objrs!tglbayar, "yyyy-mm-dd") + "'  where custid='" + Trim(M_Objrs!CustId) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   

End Sub

Private Sub ISI_DATA_POP_BP2()

'@@Report POP-BP2 [24-11-09]-- POSTGREE
Dim a As String
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
cmdsql = "select * from vwlunasmax"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

cmdsql = " select custid,name,Amountwo,agent from MGM where agent not in ('LUNAS','PULLOUT') and "
If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
cmdsql = cmdsql + " and custid in("
cmdsql = cmdsql + " select custid from vwlunasmax where datediff('month',tglbayar,now())=2)"

If Combo1(0).text <> "" And Combo1(1).text <> "" Then
  cmdsql = cmdsql + " and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
End If


M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount > 0 Then
ProgressBar1.Max = M_Objrs.RecordCount
End If
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    cmdsql = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    cmdsql = cmdsql + " ("
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId) + "', "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name) + "', "
    cmdsql = cmdsql + " " + CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo)) + ", "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent) + "', "
    cmdsql = cmdsql + " 'POP BP2') "
    M_RPTCONN.execute cmdsql

M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
    
    
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
        cmdsql = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        cmdsql = cmdsql + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
        cmdsql = cmdsql + " group by custid "
   
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
    End If
    While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set payment = " + CStr(M_Objrs!Payment) + "  where custid='" + Trim(M_Objrs!CustId) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from usertbl where  "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " userid >= '" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " userid in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
        While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set spvcode = '" + CStr(M_Objrs!SPVCODE) + "'  where agent='" + Trim(M_Objrs!Userid) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "')"
    ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' ))"
    End If
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
   While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set PaymentDate = '" + Format(M_Objrs!tglbayar, "yyyy-mm-dd") + "'  where custid='" + Trim(M_Objrs!CustId) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing

End Sub

Private Sub ISI_DATA_POP_BP3()

'@@Report POP-BP3 [24-11-09]-- POSTGREE
Dim a As String
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
cmdsql = "select * from vwlunasmax"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

cmdsql = " select custid,name,Amountwo,agent from MGM where agent not in ('LUNAS','PULLOUT') and "
If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
       cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
cmdsql = cmdsql + " and custid in("
cmdsql = cmdsql + " select custid from vwlunasmax where datediff('month',tglbayar,now())>=3)"

If Combo1(0).text <> "" And Combo1(1).text <> "" Then
  cmdsql = cmdsql + " and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
End If

M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount > 0 Then
ProgressBar1.Max = M_Objrs.RecordCount
End If
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    cmdsql = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    cmdsql = cmdsql + " ("
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId) + "', "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name) + "', "
    cmdsql = cmdsql + " " + CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo)) + ", "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent) + "', "
    cmdsql = cmdsql + " 'POP BP3') "
    M_RPTCONN.execute cmdsql

M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
    
    
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
        cmdsql = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        cmdsql = cmdsql + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
        cmdsql = cmdsql + " group by custid "
   
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
    End If
    While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set payment = " + CStr(M_Objrs!Payment) + "  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from usertbl where  "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " userid >= '" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " userid in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
        While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set spvcode = '" + CStr(M_Objrs!SPVCODE) + "'  where agent='" + Trim(M_Objrs!Userid) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "')"
    ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent in(select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode <= '" + Combo3(1).text + "'))"
    End If
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
   While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set PaymentDate = '" + Format(M_Objrs!tglbayar, "yyyy-mm-dd") + "'  where custid='" + Trim(M_Objrs!CustId) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing


End Sub

Private Sub ISI_DATA_BP1()

'@@Report BP1 [24-11-09] -- POSTGRESS
Dim a As String
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwptp1 kosong! [25-11-09]
cmdsql = "select * from vwptp1"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

cmdsql = " select custid,name,Amountwo,agent from MGM where "
If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
       cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
cmdsql = cmdsql + " and custid in("
cmdsql = cmdsql + " select custid from vwptp1 where datediff('month',promisedate::date,now()::date)=1 and custid not in (select distinct custid from tbllunas where custid<>''))"
'cmdsql = cmdsql + " select custid from vwptp1 where datediff('month',promisedate::date,now()::date)=1 and custid not in (select distinct custid from tbllunas))"

If Combo1(0).text <> "" And Combo1(1).text <> "" Then
  cmdsql = cmdsql + " and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
End If

DoEvents
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount > 0 Then
ProgressBar1.Max = M_Objrs.RecordCount
End If
While Not M_Objrs.EOF
DoEvents
    ProgressBar1.Value = M_Objrs.Bookmark
    cmdsql = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    cmdsql = cmdsql + " ("
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId) + "', "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name) + "', "
    cmdsql = cmdsql + " " + CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo)) + ", "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent) + "', "
    cmdsql = cmdsql + " 'BP1') "
    M_RPTCONN.execute cmdsql

M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
    
    
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
        cmdsql = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        cmdsql = cmdsql + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
        cmdsql = cmdsql + " group by custid "
   
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
    End If
    While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set payment = " + CStr(M_Objrs!Payment) + "  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from usertbl where  "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " userid >= '" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " userid in(select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode <= '" + Combo3(1).text + "') "
   End If
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
        While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set spvcode = '" + CStr(M_Objrs!SPVCODE) + "'  where agent='" + Trim(M_Objrs!Userid) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "')"
    ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent in(select userid from usertbl where spvcode  >= '" + Combo3(0).text + "' and spvcode <= '" + Combo3(1).text + "'))"
    End If
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
   While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set PaymentDate = '" + Format(M_Objrs!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   

End Sub

Private Sub ISI_DATA_BP2()

'@@Report BP2 [23-11-09]
Dim a As String
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwptp1 kosong! [25-11-09]
cmdsql = "select * from vwptp1"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

cmdsql = " select custid,name,Amountwo,agent from MGM where "
If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
       cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
cmdsql = cmdsql + " and custid in("
cmdsql = cmdsql + " select custid from vwptp1 where datediff('month',promisedate::date,now()::date)=2 and custid not in (select distinct custid from tbllunas where custid<>''))"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
DoEvents



If M_Objrs.RecordCount > 0 Then
ProgressBar1.Max = M_Objrs.RecordCount
End If
While Not M_Objrs.EOF
DoEvents
    ProgressBar1.Value = M_Objrs.Bookmark
    cmdsql = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    cmdsql = cmdsql + " ("
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId) + "', "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name) + "', "
    cmdsql = cmdsql + " " + CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo)) + ", "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent) + "', "
    cmdsql = cmdsql + " 'BP2') "
    M_RPTCONN.execute cmdsql

M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
    
    
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
        cmdsql = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        cmdsql = cmdsql + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
        cmdsql = cmdsql + " group by custid "
   
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
    End If
    While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set payment = " + CStr(M_Objrs!Payment) + "  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from usertbl where  "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " userid >= '" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " userid in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode <= '" + Combo3(1).text + "') "
   End If
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
        While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set spvcode = '" + CStr(M_Objrs!SPVCODE) + "'  where agent='" + Trim(M_Objrs!Userid) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "')"
    ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent in(select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and Spvcode <= '" + Combo3(1).text + "'))"
    End If
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
   While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set PaymentDate = '" + Format(M_Objrs!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   

End Sub

Private Sub ISI_DATA_BP3()

'@@Report BP3 [23-11-09]
Dim a As String
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwptp1 kosong! [25-11-09]
cmdsql = "select * from vwptp1"
M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

cmdsql = " select custid,name,Amountwo,agent from MGM where "
If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
       cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
cmdsql = cmdsql + " and custid in("
cmdsql = cmdsql + " select custid from vwptp1 where datediff('month',promisedate,now())>=3 and custid not in (select distinct custid from tbllunas))"

If Combo1(0).text <> "" And Combo1(1).text <> "" Then
  cmdsql = cmdsql + " and recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
End If

M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_Objrs.RecordCount > 0 Then
ProgressBar1.Max = M_Objrs.RecordCount
End If
While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
    cmdsql = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    cmdsql = cmdsql + " ("
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId) + "', "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name) + "', "
    cmdsql = cmdsql + " " + CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo)) + ", "
    cmdsql = cmdsql + " '" + IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent) + "', "
    cmdsql = cmdsql + " 'BP3') "
    M_RPTCONN.execute cmdsql

M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
    
    
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
        cmdsql = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        cmdsql = cmdsql + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " agent >= '" + Combo2(0).text + "' and agent <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode  <= '" + Combo3(1).text + "' )"
   End If
        cmdsql = cmdsql + " group by custid "
   
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
    End If
    While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set payment = " + CStr(M_Objrs!Payment) + "  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from usertbl where  "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " userid >= '" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' "
   ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " userid in(select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode <= '" + Combo3(1).text + "') "
   End If
   
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
        While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set spvcode = '" + CStr(M_Objrs!SPVCODE) + "'  where agent='" + Trim(M_Objrs!Userid) + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   
   
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   cmdsql = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "')"
    ElseIf Option1(1).Value = True Then
        cmdsql = cmdsql + " (select custid from mgm where agent in (select userid from usertbl where spvcode >= '" + Combo3(0).text + "' and spvcode <= '" + Combo3(1).text + "'))"
    End If
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_Objrs.RecordCount > 0 Then
    ProgressBar1.Max = M_Objrs.RecordCount
    End If
   While Not M_Objrs.EOF
    ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = " update RptPOP_BP set PaymentDate = '" + Format(M_Objrs!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_Objrs!CustId + "'"
        M_RPTCONN.execute cmdsql
   M_Objrs.MoveNext
   Wend
   Set M_Objrs = Nothing
   

End Sub
Private Sub Isi_Data_PTP_LOG()
Dim rsttemp As New ADODB.Recordset
M_RPTCONN.execute "delete from Rptnegoptp_log"
cmdsql = "select agent, sum(promisepay) as volume, substring(promisedate,1,7) as bulanbayar from tblnegoptp_log "
cmdsql = cmdsql + " where promisedate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
cmdsql = cmdsql + " and custid in(select custid from mgm where recsource between '" + Combo1(0).text + "' and '" + Combo1(1).text + "')"

If Option1(0).Value = True Then
cmdsql = cmdsql + " and agent between '" + Combo2(0).text + "' and '" + Combo2(1).text + "'  "
End If

If Option1(1).Value = True Then
cmdsql = cmdsql + " and agent in(select userid from usertbl where spvcode between '" + Combo3(0).text + "' and '" + Combo3(1).text + "')"
End If
cmdsql = cmdsql + " group by agent,bulanbayar"

Set rsttemp = New ADODB.Recordset
rsttemp.CursorLocation = adUseClient
rsttemp.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If rsttemp.RecordCount > 0 Then
ProgressBar1.Max = rsttemp.RecordCount
End If
While Not rsttemp.EOF
cmdsql = "insert into Rptnegoptp_log(agent,bulanbayar,volume)values"
cmdsql = cmdsql + "('" + IIf(IsNull(rsttemp!agent), "", rsttemp!agent) + "', "
cmdsql = cmdsql + "'" + CStr(IIf(IsNull(rsttemp!bulanbayar), Null, Format(rsttemp!bulanbayar, "yyyy-mm"))) + "', "
cmdsql = cmdsql + " " + CStr(IIf(IsNull(rsttemp!Volume), "0", rsttemp!Volume)) + ") "
M_RPTCONN.execute cmdsql
rsttemp.MoveNext
Wend
Set rsttemp = Nothing

'Exit Sub
'adderr:
'MsgBox Err.Description

End Sub
Public Sub GETPTPNEW()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
Dim jenis As String
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete  from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Check1.Value = 1 Then
jenis = "REG"
ElseIf Check2.Value = 1 Then
jenis = "IPO"

ElseIf Check3.Value = 1 Then
jenis = "RPO"
Else
'jenis = "REG"
End If

If Option1(1).Value = True Then
'CMDSQL = "Select * from  where f_cek like '%PTP%' and INputdate "

'cmdsql = " select custid,promisedate,promisepay,date(tglptpnew) AS tglptpnewnum,name,f_cek ,agent  from ("
'cmdsql = cmdsql + " SELECT MGM.custid,tblnegoptp_log.promisedate,mgm.tglptpnew,tblnegoptp_log.promisepay,mgm.recsource,mgm.agent,mgm.name,f_cek "
'cmdsql = cmdsql + " FROM tblnegoptp_log,MGM WHERE MGM.CUSTID=tblnegoptp_log.CUSTID AND TGLPTPNEW IS NOT NULL  AND DATE(tblnegoptp_log.PROMISEDATE)=DATE(MGM.TGLPTPNEW)"
'cmdsql = cmdsql + " GROUP BY MGM.CUSTID,tblnegoptp_log.PROMISEDATE,mgm.tglptpnew,tblnegoptp_log.promisepay,mgm.recsource,mgm.agent,mgm.name,f_cek) as a"
'cmdsql = cmdsql + " where  tglptpnew  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
'cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') "
'cmdsql = cmdsql + " ORDER BY AGENT"

    cmdsql = " select *,dateptpnew as PromiseDate,amountnew AS PromisePay,date(tglptpnew) AS tglptpnewnum   from MGM "
    cmdsql = cmdsql + " where  tglptpnew  "
    cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
    cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
    cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
    '@@16 APRIL 2012, TAMBAHKAN filter berdasarkan status f_cek_new='PTP-NEW'
    '@@19 Oktober 2012, Filter F_Cek_new dibuang
    'CMDSQL = CMDSQL + " and f_cek_new='PTP-NE' "
    cmdsql = cmdsql + " ORDER BY AGENT"
   M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        cmdsql = " select *,dateptpnew as promisedate,amountnew AS PromisePay ,date(tglptpnew) AS tglptpnewnum from MGM "
        cmdsql = cmdsql + " WHERE  tglptpnew between "
        cmdsql = cmdsql + "  '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent between '" + Combo2(0).text + "' and  '" + Combo2(1).text + "' "
        cmdsql = cmdsql + " ORDER BY AGENT "

'cmdsql = " select custid,promisedate,promisepay,date(tglptpnew) AS tglptpnewnum,name,f_cek ,agent from ("
'cmdsql = cmdsql + " SELECT MGM.custid,tblnegoptp_log.promisedate,mgm.tglptpnew,tblnegoptp_log.promisepay,mgm.recsource,mgm.agent,mgm.name,f_cek "
'cmdsql = cmdsql + " FROM tblnegoptp_log,MGM WHERE MGM.CUSTID=tblnegoptp_log.CUSTID AND TGLPTPNEW IS NOT NULL  AND DATE(tblnegoptp_log.PROMISEDATE)=DATE(MGM.TGLPTPNEW)"
'cmdsql = cmdsql + " GROUP BY MGM.CUSTID,TBLNEGOPTP.PROMISEDATE,mgm.tglptpnew,tblnegoptp_log.promisepay,mgm.mgm.recsource,mgm.agent,mgm.name,f_cek) as a"
'cmdsql = cmdsql + " WHERE  tglptpnew between "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' "
'cmdsql = cmdsql + " ORDER BY AGENT "

        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!TglPTPNew = CStr(IIf(IsNull(M_Objrs!tglptpnewnum), "2020-12-30", M_Objrs!tglptpnewnum))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
   ' M_OBJRS1!F_CEK = CStr(IIf(IsNull(M_OBJRS!F_CEK), "0", M_OBJRS!F_CEK))
    
    'm_objrs1!BaseOn = CStr(IIf(IsNull(m_objrs!CmbBaseOn), "", m_objrs!CmbBaseOn))
    m_objrs1!Principle = CStr(IIf(IsNull(M_Objrs!Principal), "0", M_Objrs!Principal))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

'' Update sumreserved
'CMDSQL = "select custid, sum(promisepay) from reportPTP where "
'CMDSQL = CMDSQL + " f_cek='PTP' and INputdate   Between '2009/06/01 12:00:00 AM' and  '2009/07/31 11:59:00 PM'  and  RECSOURCE Between  '-----' and 'ZZZZZ' and"
'CMDSQL = CMDSQL + " agent >= '-----' and agent <= 'ZZZZZ'"
'CMDSQL = CMDSQL + "group by custid"
'
'If Option1(1).Value = True Then
'CMDSQL = "select custid, sum(promisepay)as reservedPTP from reportPTP where f_cek like '%PTP-NE%' and INputdate "
'CMDSQL = CMDSQL + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'CMDSQL = CMDSQL + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
'CMDSQL = CMDSQL + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') GROUP BY CUSTID"
'   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'Else
'    If Option1(0).Value = True Then
'CMDSQL = "select custid, sum(promisepay)as reservedPTP from reportPTP where f_cek like '%PTP-NE%' and INputdate "
'CMDSQL = CMDSQL + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'CMDSQL = CMDSQL + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' GROUP BY CUSTID "
''
'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    End If
'End If
'Dim Lcustid As String
'Dim sumreserved As String
'While Not M_OBJRS.EOF
'Lcustid = CStr(IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId))
'sumreserved = IIf(IsNull(M_OBJRS!ReservedPTP), "0", M_OBJRS!ReservedPTP)
'M_RPTCONN.Execute "UPDATE TrackingPTP set sumreserved=" + sumreserved + " where custid='" + Trim(Lcustid) + "'"
'M_OBJRS.MoveNext
'Wend
'
'Set M_OBJRS = Nothing
Exit Sub
'


Isi_AgentErr:
  '  MsgBox Err.Description
    'Resume

End Sub
Private Sub TrackingReservedPTPversi2()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
'Dim jenis As String
On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then
    cmdsql = "select stsmove,custid,promisepay as ReservedPTP, recsource,agent,promisedate, name,amountwo,ttlptp,tenor  from reportreservenew  "
    cmdsql = cmdsql + " where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
    cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
    cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') and stsmove='0' "
    cmdsql = cmdsql + " group by custid, stsmove,promisepay ,recsource,agent,name, promisedate,AMOUNTWO,ttlptp,TENOR order by promisedate "

'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
'cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') "
'cmdsql = cmdsql + " ORDER BY AGENT"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
    
    cmdsql = "select stsmove ,custid,promisepay as  ReservedPTP, recsource,agent,promisedate, name, AMOUNTWO,ttlptp,TENOR  from reportreserve  "
cmdsql = cmdsql + " where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' AND AGENT >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "' and stsmove='0'"
cmdsql = cmdsql + " group by custid, stsmove,promisepay ,recsource,agent,name, promisedate,AMOUNTWO, ttlptp,TENOR"
    
'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' "
'cmdsql = cmdsql + " ORDER BY AGENT "
'
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!sumreserved = CStr(IIf(IsNull(M_Objrs!ReservedPTP), "0", M_Objrs!ReservedPTP))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!ttlptp), "0", M_Objrs!ttlptp))
    m_objrs1!Tenor = CStr(IIf(IsNull(M_Objrs!Tenor), "0", M_Objrs!Tenor))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1!stsmove = CStr(IIf(IsNull(M_Objrs!stsmove), "0", M_Objrs!stsmove))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub
Isi_AgentErr:
MsgBox err.Description
End Sub
Public Sub isidatabpday()
Dim m_objrs1 As New ADODB.Recordset
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then
    cmdsql = " select mgm.custid,mgm.name,promise,promisepay,agent,amountwo from mgm  inner join vwnegoptplast on mgm.custid=vwnegoptplast.custid "
    cmdsql = cmdsql + " where tglstatus between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
    cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
    cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') and f_cek='BP-'"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
    
    cmdsql = " select mgm.custid,mgm.name,promise,promisepay,agent, amountwo from mgm  inner join vwnegoptplast on mgm.custid=vwnegoptplast.custid "
    cmdsql = cmdsql + " where tglstatus between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and "
    cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
    cmdsql = cmdsql + " and AGENT >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "' and f_cek like ='BP-'"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from rptPOP_BP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
DoEvents
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!PromiseDate = CStr(IIf(IsNull(M_Objrs!promise), "2020-12-30", M_Objrs!promise))
    m_objrs1!Payment = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

End Sub
Private Sub PTP_log()
Dim M_Objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
'Dim jenis As String
'On Error GoTo Isi_AgentErr

M_RPTCONN.execute "Delete * from TrackingPTP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then
cmdsql = "select custid,promisepay as ReservedPTP,count(custid) as tenor1, recsource,agent,promisedate, name,amountwo,ttlptp,tenor  from reportptplog "
cmdsql = cmdsql + " where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in "
cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "') "
cmdsql = cmdsql + " group by custid, recsource,agent,name, promisedate,AMOUNTWO,ttlptp,TENOR,promisepay "

'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
'cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') "
'cmdsql = cmdsql + " ORDER BY AGENT"
 M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
    
        cmdsql = "select custid,promisepay as  ReservedPTP, recsource,agent,promisedate, name, AMOUNTWO,ttlptp,TENOR  from reportptplog  "
        cmdsql = cmdsql + " where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' AND AGENT >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "'"
        cmdsql = cmdsql + " group by custid, recsource,agent,name, promisedate,AMOUNTWO, ttlptp,TENOR,promisepay "
    
'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' "
'cmdsql = cmdsql + " ORDER BY AGENT "
'
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!sumreserved = CStr(IIf(IsNull(M_Objrs!ReservedPTP), "0", M_Objrs!ReservedPTP))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_Objrs!ttlptp), "0", M_Objrs!ttlptp))
    m_objrs1!Tenor = CStr(IIf(IsNull(M_Objrs!Tenor), "0", M_Objrs!Tenor))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub
Isi_AgentErr:
MsgBox err.Description
End Sub
Public Sub isidataBpMonth()
Dim m_objrs1 As New ADODB.Recordset
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then
'---sintak buat tarik bp mounth
        cmdsql = " select * ,datediff('month',date(promisedate),date(now())) as selisih  from ( "
        cmdsql = cmdsql + "select mgm.custid,mgm.name,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek,f_cek_new,mgm.recsource from  mgm "
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where  f_cek_new='BP-' and  "
        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in  "
        cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')"
        cmdsql = cmdsql + " and custid  not in (select custid from ( "
        cmdsql = cmdsql + " select mgm.custid,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek,f_cek_new from  mgm"
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where date_part('month',promisedate) = date_part('month',now()) and date_part('year',promisedate) = date_part('year',now()) and custid<>''  ) "
        cmdsql = cmdsql + " AND promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and datediff('month',date(promisedate),date(now()))>=0 "
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        cmdsql = " select * ,datediff('month',date(promisedate),date(now())) as selisih  from ( "
        cmdsql = cmdsql + "select mgm.custid,mgm.name,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek,f_cek_new,mgm.recsource from  mgm "
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where  f_cek='BP-' and  "
        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
        cmdsql = cmdsql + " and AGENT >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "'"
        cmdsql = cmdsql + " and custid  not in (select custid from ( "
        cmdsql = cmdsql + " select mgm.custid,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek,f_cek_new from  mgm"
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where date_part('month',promisedate) = date_part('month',now()) and date_part('year',promisedate) = date_part('year',now()) and custid<>'' ) "
        cmdsql = cmdsql + " AND promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and datediff('month',date(promisedate),date(now()))>=0 "
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from rptPOP_BP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
DoEvents
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!PromiseDate = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!Payment = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1!Category = CStr(IIf(IsNull(M_Objrs!selisih), "0", CStr(Val(M_Objrs!selisih))))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing


End Sub
Public Sub isidatabpday2()

Dim m_objrs1 As New ADODB.Recordset
M_RPTCONN.execute "Delete From rptPOP_BP"
Set M_Objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

If Option1(1).Value = True Then

'---sintak buat tarik bp mounth
        cmdsql = " select * ,datediff('month',date(promisedate),date(now())) as selisih  from ( "
        cmdsql = cmdsql + "select mgm.custid,mgm.name,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,,f_cek,f_cek_new,mgm.recsource from  mgm "
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where  f_cek_new='BP-' and  "
        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' and agent in  "
        cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')"
        cmdsql = cmdsql + " and custid  in (select custid from ( "
        cmdsql = cmdsql + " select mgm.custid,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek_new from  mgm"
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where date_part('month',promisedate) = date_part('month',now()) and date_part('year',promisedate) = date_part('year',now()) and custid<>'') "
        cmdsql = cmdsql + " AND promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND date(now())-date(promisedate)>7 "
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        cmdsql = " select * ,datediff('month',date(promisedate),date(now())) as selisih  from ( "
        cmdsql = cmdsql + "select mgm.custid,mgm.name,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,,f_cek,f_cek_new,mgm.recsource from  mgm "
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where  f_cek_new='BP-' and  "
        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "' "
        cmdsql = cmdsql + " and AGENT >= '" + Combo1(0).text + "' and agent <= '" + Combo1(1).text + "'"
        cmdsql = cmdsql + " and custid   in (select custid from ( "
        cmdsql = cmdsql + " select mgm.custid,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek_new from  mgm"
        cmdsql = cmdsql + " Inner Join ( "
        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
        cmdsql = cmdsql + " where date_part('month',promisedate) = date_part('month',now()) and date_part('year',promisedate) = date_part('year',now()) and custid<>'') "
        cmdsql = cmdsql + " AND promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND date(now())-date(promisedate)>7 "
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from rptPOP_BP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_Objrs.RecordCount + 1
While Not M_Objrs.EOF
ProgressBar1.Value = M_Objrs.Bookmark
DoEvents
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_Objrs!agent), "", M_Objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_Objrs!CustId), "", M_Objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_Objrs!Name), "", M_Objrs!Name))
    m_objrs1!PromiseDate = CStr(IIf(IsNull(M_Objrs!PromiseDate), "2020-12-30", M_Objrs!PromiseDate))
    m_objrs1!Payment = CStr(IIf(IsNull(M_Objrs!PromisePay), "0", M_Objrs!PromisePay))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_Objrs!amountwo), "0", M_Objrs!amountwo))
    m_objrs1!Category = CStr(IIf(IsNull(M_Objrs!selisih), "0", CStr(Val(M_Objrs!selisih))))
    m_objrs1.update
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
Set m_objrs1 = Nothing

End Sub
'Dim M_OBJRS1 As New ADODB.Recordset
'M_RPTCONN.Execute "Delete From rptPOP_BP"
'Set m_objrs = New ADODB.Recordset
'Set M_OBJRS1 = New ADODB.Recordset
'm_objrs.CursorLocation = adUseClient
'M_OBJRS1.CursorLocation = adUseClient
'
'If Option1(1).Value = True Then
''---sintak buat tarik bp mounth
'        cmdsql = " select * ,datediff('month',date(promisedate),date(now())) as selisih  from ( "
'        cmdsql = cmdsql + "select mgm.custid,mgm.name,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek,mgm.recsource from  mgm "
'        cmdsql = cmdsql + " Inner Join ( "
'        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
'        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
'        cmdsql = cmdsql + " where  f_cek='BP-' and  "
'        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in  "
'        cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')"
'        cmdsql = cmdsql + " AND promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' "
'        m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'Else
'    If Option1(0).Value = True Then
'        cmdsql = " select * ,datediff('month',date(promisedate),date(now())) as selisih  from ( "
'        cmdsql = cmdsql + "select mgm.custid,mgm.name,a.promisedate,a.promisepay,mgm.amountwo,mgm.tglstatus,mgm.agent,f_cek,mgm.recsource from  mgm "
'        cmdsql = cmdsql + " Inner Join ( "
'        cmdsql = cmdsql + " select * from tblnegoptp where id in ( "
'        cmdsql = cmdsql + " select max(id) as id  from tblnegoptp group by custid )) as a on mgm.custid=a.custid) as tbl "
'        cmdsql = cmdsql + " where  f_cek='BP-' and  "
'        cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
'        cmdsql = cmdsql + " and AGENT >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "'"
'        cmdsql = cmdsql + " AND promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' "
'        m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    End If
'End If
'M_OBJRS1.Open "Select * from rptPOP_BP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
'ProgressBar1.Max = m_objrs.RecordCount + 1
'While Not m_objrs.EOF
'ProgressBar1.Value = m_objrs.Bookmark
'DoEvents
'    M_OBJRS1.AddNew
'    M_OBJRS1!agent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
'    M_OBJRS1!CustId = CStr(IIf(IsNull(m_objrs!CustId), "", m_objrs!CustId))
'    M_OBJRS1!Name = CStr(IIf(IsNull(m_objrs!Name), "", m_objrs!Name))
'    M_OBJRS1!PromiseDate = CStr(IIf(IsNull(m_objrs!PromiseDate), "2020-12-30", m_objrs!PromiseDate))
'    M_OBJRS1!Payment = CStr(IIf(IsNull(m_objrs!PromisePay), "0", m_objrs!PromisePay))
'    M_OBJRS1!amountwo = CStr(IIf(IsNull(m_objrs!amountwo), "0", m_objrs!amountwo))
'    M_OBJRS1!category = CStr(IIf(IsNull(m_objrs!selisih), "0", CStr(Val(m_objrs!selisih))))
'    M_OBJRS1.UPDATE
'    m_objrs.MoveNext
'Wend
'Set m_objrs = Nothing
'Set M_OBJRS1 = Nothing
'
'End Sub

Private Sub report_ot_isi_previous()
    Dim M_Objrs As New ADODB.Recordset
    Dim cmdsql As String
    Dim cmdsql_update As String

    cmdsql = "select * from mgm_hst where id in "
    cmdsql = cmdsql + "(select max(id) as id from mgm_hst where agent in "
    If Option1(1).Value = True Then
        cmdsql = cmdsql + "(select userid from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "' AND USERTYPE =1) and "
    Else
        If Option1(0).Value = True Then
            cmdsql = cmdsql + "(select userid from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' AND USERTYPE =1) and "
        Else
            cmdsql = cmdsql + "(select userid from usertbl where AKTIF = 0 AND USERTYPE =1') and "
        End If
    End If
    
'    cmdsql = cmdsql + " custid in (select custid from mgm where recsource between '"
'    cmdsql = cmdsql + Combo1(0).Text + "' and '" + Combo1(1).Text + "') and date(tgl) between '"
'    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
'    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' and date_part('hour',tgl) < '17' "
'    cmdsql = cmdsql + "group by custid,agent)"
    
    cmdsql = cmdsql + " custid in (select custid from mgm where recsource between '"
    cmdsql = cmdsql + Combo1(0).text + "' and '" + Combo1(1).text + "') "
    cmdsql = cmdsql + "and date_part('hour',tgl) < '17' and f_cek_new is not null "
    cmdsql = cmdsql + "group by custid,agent)"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql_update = "update reportagenttelp set status_call_previous='"
        cmdsql_update = cmdsql_update + IIf(IsNull(M_Objrs("statuscall")), "", M_Objrs("statuscall")) + "',status_ptp_previous='"
        cmdsql_update = cmdsql_update + IIf(IsNull(M_Objrs("f_cek")), "", M_Objrs("f_cek")) + "',tglcall_previous='"
        cmdsql_update = cmdsql_update + CStr(IIf(IsNull(M_Objrs("tgl")), "", M_Objrs("tgl"))) + "' where custid='"
        cmdsql_update = cmdsql_update + M_Objrs("custid") + "' and agent='"
        cmdsql_update = cmdsql_update + M_Objrs("agent") + "'"
        M_RPTCONN.execute cmdsql_update
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    Exit Sub
Isi_AgentErr:
        MsgBox err.Description
        'Resume
End Sub

Private Sub report_ot_agent()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim cmdsql_insert As String
    
    M_RPTCONN.execute "delete from reportagenttelp"
    
    cmdsql = "select f_cek_new,statuscall,agent,custid,tglcall from mgm where agent in "
    If Option1(1).Value = True Then
        cmdsql = cmdsql + "(select userid from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "' AND USERTYPE =1) and "
    Else
        If Option1(0).Value = True Then
            cmdsql = cmdsql + "(select userid from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).text + "' and userid <= '" + Combo2(1).text + "' AND USERTYPE =1) and "
        Else
            cmdsql = cmdsql + "(select userid from usertbl where AKTIF = 0 AND USERTYPE =1') and "
        End If
    End If
    
    cmdsql = cmdsql + "recsource between '"
    cmdsql = cmdsql + Combo1(0).text + "' and '" + Combo1(1).text + "' and date(tglcall) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' and date_part('hour',tglcall) >='17'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ProgressBar1.Max = M_Objrs.RecordCount + 2
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql_insert = "insert into  reportagenttelp (agent,custid,tglcall_current,"
            cmdsql_insert = cmdsql_insert + "status_call_current,status_ptp_current) values ('"
            cmdsql_insert = cmdsql_insert + M_Objrs("agent") + "','"
            cmdsql_insert = cmdsql_insert + M_Objrs("custid") + "','"
            cmdsql_insert = cmdsql_insert + CStr(M_Objrs("tglcall")) + "','"
            cmdsql_insert = cmdsql_insert + M_Objrs("statuscall") + "','"
            cmdsql_insert = cmdsql_insert + M_Objrs("f_cek_new") + "')"
            M_RPTCONN.execute cmdsql_insert
        M_Objrs.MoveNext
    Wend
End Sub


'================================================================================================
'@@ 17-03-2011 Report Contactto
Private Sub IsiAgentContactto()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    If Option1(1).Value = True Then
        cmdsql = "select distinct u.spvcode as spv ,m.agent as agent"
        cmdsql = cmdsql + " from mgm as m, usertbl as u where "
        cmdsql = cmdsql + " m.agent=u.userid and u.userid in (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and m.tglcall between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' and "
        cmdsql = cmdsql + " m.recsource between '"
        cmdsql = cmdsql + Trim(Combo1(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo1(1).text) + "' "
        cmdsql = cmdsql + " group by m.agent,u.spvcode "
        cmdsql = cmdsql + "order by u.spvcode,m.agent asc"
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    M_RPTCONN.execute "delete from TblRptContactto"
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into TblRptContactto (spvcode,agent) values ('"
            cmdsql = cmdsql + Trim(M_Objrs("spv")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("agent")) + "')"
             M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub

Private Sub IsiContactto()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
'    If Option1(1).Value = True Then
'        CMDSQL = "select u.spvcode as spv ,m.agent as agent,m.stscallwith as status,count(m.stscallwith) as jumlah "
'        CMDSQL = CMDSQL + " from mgm as m, usertbl as u where "
'        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
'        CMDSQL = CMDSQL + " where spvcode between '"
'        CMDSQL = CMDSQL + Trim(Combo3(0).Text) + "' and '"
'        CMDSQL = CMDSQL + Trim(Combo3(1).Text) + "' and usertype='1') and m.tglcall between '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' and "
'        CMDSQL = CMDSQL + " m.recsource between '"
'        CMDSQL = CMDSQL + Trim(Combo1(0).Text) + "' and '"
'        CMDSQL = CMDSQL + Trim(Combo1(1).Text) + "' "
'        CMDSQL = CMDSQL + " group by m.agent,m.stscallwith,u.spvcode "
'        CMDSQL = CMDSQL + "order by m.agent,m.stscallwith,u.spvcode asc"
'    End If
    
    '@@ 01/06/2011, diubah sintak SQLnya ambil dari MGM_HST aja
    If Option1(1).Value = True Then
        cmdsql = "select u.spvcode as spv ,m.agent as agent,m.ststelpwith as status,count(m.ststelpwith) as jumlah "
        cmdsql = cmdsql + " from mgm_hst as m, usertbl as u where "
        cmdsql = cmdsql + " m.agent=u.userid and u.userid in (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and m.tgl between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' and "
        cmdsql = cmdsql + " m.ststelpwith in ('OTHER','CH','SPOUSE','PARENT')"
'        CMDSQL = CMDSQL + " m.recsource between '"
'        CMDSQL = CMDSQL + Trim(Combo1(0).Text) + "' and '"
'        CMDSQL = CMDSQL + Trim(Combo1(1).Text) + "' "
        cmdsql = cmdsql + " group by m.agent,m.ststelpwith,u.spvcode "
        cmdsql = cmdsql + "order by m.agent,m.ststelpwith,u.spvcode asc"
    End If
    
    
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
'             If Trim(m_objrs("status")) <> "SPOUSE" Or _
'                Trim(m_objrs("status")) <> "CONTACTED-CH" Or _
'                Trim(m_objrs("status")) <> "OTHER" Or _
'                Trim(m_objrs("status")) <> "PARENT" Or _
'                Trim(m_objrs("status")) <> "CH" Or _
'                Trim(m_objrs("status")) = "" Then
'                m_objrs.MoveNext
'             End If
            On Error Resume Next
            ProgressBar1.Value = M_Objrs.Bookmark
             cmdsql = "update tblrptcontactto set ["
             cmdsql = cmdsql + Trim(Replace(M_Objrs("status"), "/", "")) + "]='"
             cmdsql = cmdsql + CStr(M_Objrs("jumlah")) + "' where spvcode='"
             cmdsql = cmdsql + Trim(M_Objrs("spv")) + "' and agent='"
             cmdsql = cmdsql + Trim(M_Objrs("agent")) + "'"
             M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub
'@@ 01 Juni 2011
Private Sub IsiContacttoJmlAcc()
    Dim M_Objrs As ADODB.Recordset
    Dim m_objrs_rpt As ADODB.Recordset
    Dim cmdsql As String
    
     '@@ 01/06/2011, diubah sintak SQLnya ambil dari MGM_HST aja
'    If Option1(1).Value = True Then
'        CMDSQL = "select u.spvcode as spv ,m.agent as agent,count(m.agent) as jumlah "
'        CMDSQL = CMDSQL + " from mgm_hst as m, usertbl as u where "
'        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
'        CMDSQL = CMDSQL + " where spvcode between '"
'        CMDSQL = CMDSQL + Trim(Combo3(0).Text) + "' and '"
'        CMDSQL = CMDSQL + Trim(Combo3(1).Text) + "' and usertype='1') and m.tgl between '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "'  "
'        'CMDSQL = CMDSQL + " m.ststelpwith in ('OTHER','CH','SPOUSE','OTHER')"
''        CMDSQL = CMDSQL + " m.recsource between '"
''        CMDSQL = CMDSQL + Trim(Combo1(0).Text) + "' and '"
''        CMDSQL = CMDSQL + Trim(Combo1(1).Text) + "' "
'        CMDSQL = CMDSQL + " group by m.agent,u.spvcode "
'        CMDSQL = CMDSQL + "order by m.agent,u.spvcode asc"
'    End If
    
    
    'Ambil Data agentnya
    cmdsql = "select spvcode,agent from tblrptcontactto order by spvcode,agent"
    Set m_objrs_rpt = New ADODB.Recordset
    m_objrs_rpt.CursorLocation = adUseClient
    m_objrs_rpt.Open cmdsql, M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs_rpt.RecordCount > 0 Then
        ProgressBar1.Max = m_objrs_rpt.RecordCount
        While Not m_objrs_rpt.EOF
            ProgressBar1.Value = m_objrs_rpt.Bookmark
            cmdsql = "select distinct custid from mgm_hst where agent='"
            cmdsql = cmdsql + Trim(m_objrs_rpt("agent")) + "' and tgl between '"
            cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
            cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "'  "
            
            
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            'update data ke access
            cmdsql = "update tblrptcontactto set jml_acc='"
            cmdsql = cmdsql + CStr(M_Objrs.RecordCount) + "' where agent='"
            cmdsql = cmdsql + Trim(m_objrs_rpt("agent")) + "' and spvcode='"
            cmdsql = cmdsql + Trim(m_objrs_rpt("spvcode")) + "'"
            M_RPTCONN.execute cmdsql
            
            Set M_Objrs = Nothing
            
            m_objrs_rpt.MoveNext
        Wend
    End If
    
    Set m_objrs_rpt = Nothing
    
'
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_OBJRS.RecordCount > 0 Then
'        ProgressBar1.Max = M_OBJRS.RecordCount
'        While Not M_OBJRS.EOF
'            On Error Resume Next
'            ProgressBar1.Value = M_OBJRS.Bookmark
''             CMDSQL = "update tblrptcontactto set ["
''             CMDSQL = CMDSQL + Trim(Replace(M_OBJRS("status"), "/", "")) + "]='"
''             CMDSQL = CMDSQL + CStr(M_OBJRS("jumlah")) + "' where spvcode='"
''             CMDSQL = CMDSQL + Trim(M_OBJRS("spv")) + "' and agent='"
''             CMDSQL = CMDSQL + Trim(M_OBJRS("agent")) + "'"
'
'             CMDSQL = "update tblrptcontactto set jml_acc='"
'             CMDSQL = CMDSQL + CStr(M_OBJRS("jumlah")) + "' where spvcode='"
'             CMDSQL = CMDSQL + Trim(M_OBJRS("spv")) + "' and agent='"
'             CMDSQL = CMDSQL + Trim(M_OBJRS("agent")) + "'"
'             M_RPTCONN.Execute CMDSQL
'            M_OBJRS.MoveNext
'        Wend
'    End If
'    Set M_OBJRS = Nothing
End Sub




'@@ 05 Mei 2011 Report Request BS
Private Sub Isi_Bs()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    If Option1(0).Value Then
        cmdsql = "select * from tbl_req_bs where agent between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and tgl_req_bs between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    Else
        cmdsql = "select * from tbl_req_bs where agent in "
        cmdsql = cmdsql + " (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and tgl_req_bs between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_RPTCONN.execute "delete from rptreqbs"
    
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rptreqbs (tgl,agent,custid,nama,year_bs,month_bs) values ('"
            cmdsql = cmdsql + CStr(Format(M_Objrs("tgl_req_bs"), "yyyy-mm-dd")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("agent")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("nama")) + "','"
            cmdsql = cmdsql + CStr(M_Objrs("year_bs")) + "','"
            cmdsql = cmdsql + CStr(M_Objrs("month_bs")) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub


'@@ 05 Mei 2011 Report Request EC
Private Sub Isi_EC()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    If Option1(0).Value Then
        cmdsql = "select * from tbl_req_ec where agent between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and tgl_req_ec between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    Else
        cmdsql = "select * from tbl_req_ec where agent in "
        cmdsql = cmdsql + " (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and tgl_req_ec between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_RPTCONN.execute "delete from rptreqec"
    
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rptreqec (tgl,agent,custid,nama) values ('"
            cmdsql = cmdsql + CStr(Format(M_Objrs("tgl_req_ec"), "yyyy-mm-dd")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("agent")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("nama")) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub

'@@ 05 Mei 2011 Report Request OST
Private Sub Isi_OST()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    If Option1(0).Value Then
        cmdsql = "select * from tbl_req_ost where agent between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and tgl_req_ost between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    Else
        cmdsql = "select * from tbl_req_ost where agent in "
        cmdsql = cmdsql + " (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and tgl_req_ost between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_RPTCONN.execute "delete from rptreqost"
    
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rptreqost (tgl,agent,custid,addr) values ('"
            cmdsql = cmdsql + CStr(Format(M_Objrs("tgl_req_ost"), "yyyy-mm-dd")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("agent")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("addr")) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub

'@@ 05 Mei 2011 Report Problem
Private Sub Isi_Problem()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    If Option1(0).Value Then
        cmdsql = "select * from tbl_req_problem where agent between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and tgl between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    Else
        cmdsql = "select * from tbl_req_problem where agent in "
        cmdsql = cmdsql + " (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and tgl between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_RPTCONN.execute "delete from rptreqproblem"
    
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rptreqproblem (tgl,agent,custid,problem,solve,nama_agent) values ('"
            cmdsql = cmdsql + CStr(Format(M_Objrs("tgl"), "yyyy-mm-dd")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("agent")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("problem")) + "','"
            cmdsql = cmdsql + Trim(IIf(IsNull(M_Objrs("solve")), "", M_Objrs("solve"))) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("nama_agent")) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub

'@@ 05 Mei 2011 Report Request PUM
Private Sub Isi_PUM()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    If Option1(0).Value Then
        cmdsql = "select * from tbl_req_pum where agent between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and tgl_req between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    Else
        cmdsql = "select * from tbl_req_pum where agent in "
        cmdsql = cmdsql + " (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and tgl_req between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_RPTCONN.execute "delete from rptreqpum"
    
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rptreqpum (tgl,agent,custid,amountwo,payment_date) values ('"
            cmdsql = cmdsql + CStr(Format(M_Objrs("tgl_req"), "yyyy-mm-dd")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("agent")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("amountwo")), "0", M_Objrs("amountwo"))) + "',"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("payment_date")), "null", "'" + Format(M_Objrs("payment_date"), "yyyy-mm-dd") + "'")) + ")"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub




'@@ 05 Mei 2011 Report Request RS
Private Sub Isi_RS()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    If Option1(0).Value Then
        cmdsql = "select * from tbl_req_rs where agent between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and tgl_req_rs between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    Else
        cmdsql = "select * from tbl_req_rs where agent in "
        cmdsql = cmdsql + " (select userid from usertbl "
        cmdsql = cmdsql + " where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1') and tgl_req_rs between '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
        cmdsql = cmdsql + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_RPTCONN.execute "delete from rptreqrs"
    
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rptreqrs (tgl,agent,custid,tot_payment,installment) values ('"
            cmdsql = cmdsql + CStr(Format(M_Objrs("tgl_req_rs"), "yyyy-mm-dd")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("agent")) + "','"
            cmdsql = cmdsql + Trim(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("tot_payment")), "0", M_Objrs("tot_payment"))) + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("installment_period")), "", M_Objrs("installment_period"))) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub

'@@ 15-03-2012 Report Hot Prospect
Private Sub Isi_Hot_Prospect()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmdsql = "select * from mgm where status_htc='1' and "
    If Option1(0).Value Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where userid between '"
        cmdsql = cmdsql + IIf(IsNull(Combo2(0).text), "", Trim(Combo2(0).text)) + "' and '"
        cmdsql = cmdsql + IIf(IsNull(Combo2(1).text), "", Trim(Combo2(1).text)) + "' "
        cmdsql = cmdsql + " and aktif='0' and usertype='1') "
    Else
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + IIf(IsNull(Combo3(0).text), "", Trim(Combo3(0).text)) + "' and '"
        cmdsql = cmdsql + IIf(IsNull(Combo3(1).text), "", Trim(Combo3(1).text)) + "' "
        cmdsql = cmdsql + " and aktif='0' and usertype='1') "
    End If
    
    cmdsql = cmdsql + " and date(tglcall) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "'"
    cmdsql = cmdsql + " order by agent,custid asc "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    M_RPTCONN.execute "delete from rpthotprospect "
    
    ProgressBar1.Max = M_Objrs.RecordCount
    Dim StatusKeep As String
    While Not M_Objrs.EOF
        On Error GoTo SALAH
        ProgressBar1.Value = M_Objrs.Bookmark
        
        If IsNull(M_Objrs("status_keep")) = True Then
            StatusKeep = ""
        Else
            StatusKeep = "Kept"
        End If
        
        cmdsql = "insert into RptHotProspect (custid,nama,agent,last_call,status_kept) values ('"
        cmdsql = cmdsql + Trim(IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))) + "','"
        cmdsql = cmdsql + Trim(IIf(IsNull(M_Objrs("name")), "", M_Objrs("name"))) + "','"
        cmdsql = cmdsql + Trim(IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))) + "','"
        cmdsql = cmdsql + CStr(Format(M_Objrs("tglcall"), "yyyy-mm-dd")) + "','"
        cmdsql = cmdsql + StatusKeep + "')"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
    Exit Sub
SALAH:
    Set M_Objrs = Nothing
    MsgBox "Ada error: " & err.Description
End Sub


'@@ 30-03-2012 Report Keep Account
Private Sub Isi_Keep_Account()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    M_RPTCONN.execute "delete from rptkeepacc "
    
    cmdsql = "select * from tbl_keep_acc where "
    If Option1(0).Value Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where userid between '"
        cmdsql = cmdsql + IIf(IsNull(Combo2(0).text), "", Trim(Combo2(0).text)) + "' and '"
        cmdsql = cmdsql + IIf(IsNull(Combo2(1).text), "", Trim(Combo2(1).text)) + "' "
        'CMDSQL = CMDSQL + " and aktif='0' and usertype='1') "
        '@@05-09-2012 usertype keseluruhan
        cmdsql = cmdsql + " and aktif='0') "
    Else
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + IIf(IsNull(Combo3(0).text), "", Trim(Combo3(0).text)) + "' and '"
        cmdsql = cmdsql + IIf(IsNull(Combo3(1).text), "", Trim(Combo3(1).text)) + "' "
        'CMDSQL = CMDSQL + " and aktif='0' and usertype='1') "
        '@@05-09-2012 usertype keseluruhan
        cmdsql = cmdsql + " and aktif='0') "
    End If
    
    cmdsql = cmdsql + " and date(tglkeep) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "'"
    cmdsql = cmdsql + " order by agent,custid asc "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    
    
    ProgressBar1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        On Error GoTo SALAH
        ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = "insert into Rptkeepacc (custid,nama,agent,tglkeep) values ('"
        cmdsql = cmdsql + Trim(IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))) + "','"
        cmdsql = cmdsql + Trim(IIf(IsNull(M_Objrs("nama")), "", M_Objrs("nama"))) + "','"
        cmdsql = cmdsql + Trim(IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))) + "','"
        cmdsql = cmdsql + CStr(Format(M_Objrs("tglkeep"), "yyyy-mm-dd")) + "')"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
    Exit Sub
SALAH:
    Set M_Objrs = Nothing
    MsgBox "Ada error: " & err.Description
End Sub

'@@ 14-05-2012 Report Unvalid Number
Private Sub Isi_Unvalid_Number()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
    cmdsql = "select * from tblunvalid_number where date(tglinput) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' and "
    If Option1(0).Value Then
        cmdsql = cmdsql + " userid between '"
        cmdsql = cmdsql + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "' "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " userid in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    
    M_RPTCONN.execute "delete from rptunvalidnumber "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    ProgressBar1.Max = M_Objrs.RecordCount
    
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = "insert into rptunvalidnumber (custid,agent,nama_agent,no_telp,"
        cmdsql = cmdsql + "status,tgl_input,keterangan) values ('"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("custid")), "", CStr(M_Objrs("custid"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("userid")), "", CStr(M_Objrs("userid"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("userinput")), "", CStr(M_Objrs("userinput"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("no_telp")), "", CStr(M_Objrs("no_telp"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("status")), "", CStr(M_Objrs("status"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("tglinput")), "", Format(M_Objrs("tglinput"), "yyyy-mm-dd")) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("keterangan")), "", M_Objrs("keterangan")) + "')"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub Isi_Valid_Number()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
    cmdsql = "select * from tblvalidnumber where date(tglinput) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' and "
    If Option1(0).Value Then
        cmdsql = cmdsql + " agent between '"
        cmdsql = cmdsql + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "' "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    
    M_RPTCONN.execute "delete from rptvalidnumber "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    ProgressBar1.Max = M_Objrs.RecordCount
    
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = "insert into rptvalidnumber (custid,agent,nama_agent,no_telp,"
        cmdsql = cmdsql + "tgl_input,keterangan) values ('"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("custid")), "", CStr(M_Objrs("custid"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("agent")), "", CStr(M_Objrs("agent"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("userinput")), "", CStr(M_Objrs("userinput"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("no_telp")), "", CStr(M_Objrs("no_telp"))) + "','"
        
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("tglinput")), "", Format(M_Objrs("tglinput"), "yyyy-mm-dd")) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("keterangan")), "", M_Objrs("keterangan")) + "')"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub RptAccReview()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
    cmdsql = "select * from tbl_log_acc_review where date(tgl) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' and "
    If Option1(0).Value Then
        cmdsql = cmdsql + " agent between '"
        cmdsql = cmdsql + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "' "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    
    M_RPTCONN.execute "delete from rptreviewacc "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    ProgressBar1.Max = M_Objrs.RecordCount
    
    While Not M_Objrs.EOF
        ProgressBar1.Value = M_Objrs.Bookmark
        cmdsql = "insert into rptreviewacc (custid,agent,no_telp,"
        cmdsql = cmdsql + "tglinput) values ('"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("custid")), "", CStr(M_Objrs("custid"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("agent")), "", CStr(M_Objrs("agent"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("telp")), "", CStr(M_Objrs("telp"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("tgl")), "", Format(M_Objrs("tgl"), "yyyy-mm-dd")) + "')"
        M_RPTCONN.execute cmdsql
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub


'@@21-06-2012, Bikin Report Log Approval SendPTP
Private Sub LogApprovalSendPTP()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    M_RPTCONN.execute "delete from rpt_log_sendptp "
    
    'CMDSQL = "select * from tblsendptp_log_approve where date(tgldata) between '"
    '@@ 25072012, Ganti nyarinya berdasarkan tanggal proposal
    cmdsql = "select * from tblsendptp_log_approve where date(tgl_proposal) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and userid between '" + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and spvcode between '" + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    
    cmdsql = cmdsql + " and tgl_proposal is not null "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rpt_log_sendptp (jenis_ptp,tgl_send_ptp,"
            cmdsql = cmdsql + "custid,date_payment_effective,"
            cmdsql = cmdsql + "tenor,pembayaran_via,tgl_tagih,agent,"
            cmdsql = cmdsql + "payment_month,total_amount_deal "
            
            cmdsql = cmdsql + ",tgl_proposal,approve_by"
            
            cmdsql = cmdsql + ") values ('"
            cmdsql = cmdsql + M_Objrs("jenis_ptp") + "','"
            cmdsql = cmdsql + Format(M_Objrs("tgldata"), "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + CStr(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + CStr(M_Objrs("tenor")) + "','"
            cmdsql = cmdsql + M_Objrs("pembayaran_via") + "','"
            cmdsql = cmdsql + Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + M_Objrs("agent") + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("payment_after_tenor")), "0", M_Objrs("payment_after_tenor"))) + "','"
            cmdsql = cmdsql + CStr(M_Objrs("total_amount_deal")) + "','"
            
            cmdsql = cmdsql + Format(M_Objrs("tgl_proposal"), "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + CStr(Trim(IIf(IsNull(M_Objrs("approve_by")), "", M_Objrs("approve_by")))) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
   End If
   
End Sub

Private Sub LogRejectedSendPTP()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    M_RPTCONN.execute "delete from rpt_log_sendptp "
    
    cmdsql = "select * from tblsendptp_log_reject where date(tgldata) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and userid between '" + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and spvcode between '" + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
   If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "insert into rpt_log_sendptp (jenis_ptp,tgl_send_ptp,"
            cmdsql = cmdsql + "custid,date_payment_effective,"
            cmdsql = cmdsql + "tenor,pembayaran_via,tgl_tagih,agent,"
            cmdsql = cmdsql + "payment_month,total_amount_deal"
            
            cmdsql = cmdsql + ",tgl_proposal,approve_by, keterangan_reject"
            
            cmdsql = cmdsql + ") values ('"
            cmdsql = cmdsql + M_Objrs("jenis_ptp") + "','"
            cmdsql = cmdsql + Format(M_Objrs("tgldata"), "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + CStr(M_Objrs("custid")) + "','"
            cmdsql = cmdsql + Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + CStr(M_Objrs("tenor")) + "','"
            cmdsql = cmdsql + M_Objrs("pembayaran_via") + "','"
            cmdsql = cmdsql + Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + M_Objrs("agent") + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("payment_after_tenor")), "0", M_Objrs("payment_after_tenor"))) + "','"
            cmdsql = cmdsql + CStr(M_Objrs("total_amount_deal")) + "',"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs("tgl_proposal")), "null", "'" + Format(M_Objrs("tgl_proposal"), "yyyy-mm-dd") + "'") + ",'"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs("log_approve")), "", M_Objrs("log_approve")) + "','"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust")) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
   End If
   
End Sub

Private Sub ContactRateLPD1()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    M_RPTCONN.execute "delete from rpt_contactrate_lpd  "
    
    cmdsql = "select * from mgm where date(tglcall) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where userid between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and usertype='1' and aktif='0' ) "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and usertype='1' and aktif='0' ) "
    End If
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = " insert into rpt_contactrate_lpd (custid,nama,status,agent) values ('"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))) + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("name")), "", M_Objrs("name"))) + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("f_contact_rate")), "", M_Objrs("f_contact_rate"))) + "','"
            cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub

'@@Report PAID OFF
Private Sub IsiCustidPaidOff()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim M_Objrs_CPA As ADODB.Recordset
    Dim RefNo As String
    Dim CustId As String
    
    M_RPTCONN.execute "delete from rptpaidoff "
    
'    CMDSQL = "select custid,tgl from mgm_hst where id in ( "
'    CMDSQL = CMDSQL + "select min(id) from mgm_hst where kodeds like '%PAID OFF%' "
'    CMDSQL = CMDSQL + " and date(tgl) between '"
'    CMDSQL = CMDSQL + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
'    If Option1(0).Value Then
'        CMDSQL = CMDSQL + " and agent in (select userid from usertbl where userid between '"
'        CMDSQL = CMDSQL + Trim(Combo2(0).Text) + "' and '"
'        CMDSQL = CMDSQL + Trim(Combo2(1).Text) + "' and aktif='0' and usertype='1' )  "
'    End If
'    If Option1(1).Value Then
'        CMDSQL = CMDSQL + " and agent in (select userid from usertbl where spvcode between '"
'        CMDSQL = CMDSQL + Trim(Combo3(0).Text) + "' and '"
'        CMDSQL = CMDSQL + Trim(Combo3(1).Text) + "' and aktif='0' and usertype='1' )  "
'    End If
'    CMDSQL = CMDSQL + " group by custid ) "
'    CMDSQL = CMDSQL + " and custid in (select custid from mgm where f_cek_new='PO-') "


    cmdsql = "select * from ("
    
    cmdsql = cmdsql + "select  custid as custid_hst,tgl,agent as agent_new from mgm_hst where id in ( "
    cmdsql = cmdsql + "select min(id) from mgm_hst where kodeds like '%PAID OFF%' "
    cmdsql = cmdsql + " and date(tgl) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where userid between '"
        cmdsql = cmdsql + Trim(Combo2(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo2(1).text) + "' and aktif='0' and usertype='1' )  "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + Trim(Combo3(0).text) + "' and '"
        cmdsql = cmdsql + Trim(Combo3(1).text) + "' and aktif='0' and usertype='1' )  "
    End If
    cmdsql = cmdsql + " group by custid ) "
    
    cmdsql = cmdsql + ") as b,mgm "
    cmdsql = cmdsql + " where mgm.custid = b.custid_hst and "
    cmdsql = cmdsql + " mgm.f_cek_new='PO-' "
    
    'CMDSQL = CMDSQL + " and mgm.custid=tblcpa.vcustid "
    'CMDSQL = CMDSQL + " and date(tblcpa.dpropsal) between '"
    'CMDSQL = CMDSQL + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    'CMDSQL = CMDSQL + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            
            'Buat Nambahin 000 untuk Card dan 0000000 Untuk PIL
            If Trim(UCase(M_Objrs("acc_type"))) = "CARD" Or _
               IsNull(M_Objrs("acc_type")) = True Or _
               M_Objrs("acc_type") = "" Then
                CustId = "000" & M_Objrs("custid")
            End If
            If Trim(UCase(M_Objrs("acc_type"))) = "PIL" Or _
               Trim(UCase(M_Objrs("acc_type"))) = "GRF" Then
                CustId = "0000000" & M_Objrs("custid")
            End If
            
            'Cari CPA Terakhirnya
            cmdsql = "select * from tblcpa where nid in ("
            cmdsql = cmdsql + "select max(nid) from tblcpa where vcustid='"
            cmdsql = cmdsql + CStr(Trim(M_Objrs("custid"))) + "')"
            Set M_Objrs_CPA = New ADODB.Recordset
            M_Objrs_CPA.CursorLocation = adUseClient
            M_Objrs_CPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs_CPA.RecordCount > 0 Then
                'Cari Refno
                If IsNull(M_Objrs_CPA("nttlpayment")) = False And _
                   IsNull(M_Objrs_CPA("nbalance")) = False Then
                    If Val(M_Objrs_CPA("nttlpayment")) < Val(M_Objrs_CPA("nbalance")) Then
                        RefNo = "D"
                    End If
                    If Val(M_Objrs_CPA("nttlpayment")) = Val(M_Objrs_CPA("nbalance")) Then
                        RefNo = "X"
'                    Else
'                        RefNo = ""
                    End If
                Else
                    RefNo = ""
                End If
                
                cmdsql = "insert into rptpaidoff (custid,tgl_paid_off,ref_no,"
                cmdsql = cmdsql + "balance,payment,payment_period,principal,"
                cmdsql = cmdsql + "payment_handle,occupation,reason,responeable_collector,justification"
                cmdsql = cmdsql + ") values ('"
                cmdsql = cmdsql + CStr(CustId) + "','"
                cmdsql = cmdsql + Format(M_Objrs("tgl"), "yyyy-mm-dd") + "','"
                'CMDSQL = CMDSQL + Format(M_OBJRS("dpropsal"), "yyyy-mm-dd") + "','"
                cmdsql = cmdsql + RefNo + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("nbalance")), "0", M_Objrs_CPA("nbalance"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("nttlpayment")), "0", M_Objrs_CPA("nttlpayment"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("nperiod")), "1", M_Objrs_CPA("nperiod"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("nprincipal")), "0", M_Objrs_CPA("nprincipal"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("vpaymenthandle")), "", M_Objrs_CPA("vpaymenthandle"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("voccupation")), "", M_Objrs_CPA("voccupation"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("vreason")), "", M_Objrs_CPA("vreason"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("agent_new")), "", M_Objrs("agent_new"))) + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs_CPA("vjust")), "", M_Objrs_CPA("vjust"))) + "')"
                M_RPTCONN.execute cmdsql
            Else
                cmdsql = "insert into rptpaidoff (custid,tgl_paid_off,responeable_collector) values ('"
                cmdsql = cmdsql + CStr(Trim(CustId)) + "','"
                cmdsql = cmdsql + Format(M_Objrs("tgl"), "yyyy-mm-dd") + "','"
                cmdsql = cmdsql + CStr(IIf(IsNull(M_Objrs("agent_new")), "", M_Objrs("agent_new"))) + "')"
                M_RPTCONN.execute cmdsql
            End If
            
            Set M_Objrs_CPA = Nothing
            M_Objrs.MoveNext
        Wend
    End If
End Sub
'------------------------------------------------ @@ 20072102 Report Call Duration LPD -----------------
Private Sub HitungDurasiLPDIcentra4()
    Dim cmdsql As String
    Dim StrKoneksi As String
    Dim M_Objrs_RitCard As ADODB.Recordset
    Dim M_Objrs_Centra As ADODB.Recordset
    
    Dim connIcentra As New ADODB.Connection
    
    M_RPTCONN.execute "delete from rptlpddurcall "
    
    '---------------------------------- Cari Yang Di Server 4 Dulu ---------------------------------------
    Set connIcentra = New ADODB.Connection
    'Server Icentra 4
    StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
    connIcentra.Open StrKoneksi
    '---------------------------------------------------------------------------------------------------

    cmdsql = "select custid,agent,f_contact_rate from mgm where date(tglcall) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "'  "
    
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where userid between '"
        cmdsql = cmdsql + CStr(Trim(Combo2(0).text)) + "' and '"
        cmdsql = cmdsql + CStr(Trim(Combo2(1).text)) + "' and usertype='1' abd aktif='0') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + CStr(Trim(Combo3(0).text)) + "' and '"
        cmdsql = cmdsql + CStr(Trim(Combo3(1).text)) + "' and usertype='1' and aktif='0') "
    End If
    
    Set M_Objrs_RitCard = New ADODB.Recordset
    M_Objrs_RitCard.CursorLocation = adUseClient
    M_Objrs_RitCard.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_Objrs_RitCard.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs_RitCard.RecordCount
        While Not M_Objrs_RitCard.EOF
            ProgressBar1.Value = M_Objrs_RitCard.Bookmark
            'Inputin Data Ke Tabel RptLPDDurCall di Access
            cmdsql = "insert into RptLPDDurCall (custid,agent) values ('"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("custid"))) + "','"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("agent"))) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs_RitCard.MoveNext
        Wend
        
        'Masukin Deh Durasi Call Yang LPD 1 dan > LPD1 diambil dari Icentra 4
        M_Objrs_RitCard.MoveFirst
        ProgressBar1.Max = M_Objrs_RitCard.RecordCount
        While Not M_Objrs_RitCard.EOF
            ProgressBar1.Value = M_Objrs_RitCard.Bookmark
            
            cmdsql = "select custid,agent,count(dur) as jmlh from report_total_working "
            cmdsql = cmdsql + " where custid='"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("custid"))) + "' and agent='"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("agent"))) + "' and "
            cmdsql = cmdsql + " date(calldate) between '"
            cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
            cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "'  "
            cmdsql = cmdsql + " group by custid,agent "
            Set M_Objrs_Centra = New ADODB.Recordset
            M_Objrs_Centra.CursorLocation = adUseClient
            M_Objrs_Centra.Open cmdsql, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_Centra.RecordCount > 0 Then
                cmdsql = " update rptlpddurcall "
                If UCase(Trim(M_Objrs_RitCard("f_contact_rate"))) = "CONTACT LPD" Then
                    cmdsql = cmdsql + " set dur_lpd1=[dur_lpd1]+"
                    cmdsql = cmdsql + CStr(M_Objrs_Centra("jmlh")) + " "
                    cmdsql = cmdsql + " where custid='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("custid")) + "' and agent='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("agent")) + "'"
                    M_RPTCONN.execute cmdsql
                End If
                If UCase(Trim(M_Objrs_RitCard("f_contact_rate"))) = "CONTACT > LPD 1" Then
                    cmdsql = cmdsql + " set dur_lpd1besar=[dur_lpd1besar]+"
                    cmdsql = cmdsql + CStr(M_Objrs_Centra("jmlh")) + " "
                    cmdsql = cmdsql + " where custid='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("custid")) + "' and agent='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("agent")) + "'"
                    M_RPTCONN.execute cmdsql
                End If
            End If
            Set M_Objrs_Centra = Nothing
            M_Objrs_RitCard.MoveNext
        Wend
        
    End If
    
    Set M_Objrs_RitCard = Nothing
    Set connIcentra = Nothing
End Sub

Private Sub HitungDurasiLPDIcentra5()
    Dim cmdsql As String
    Dim M_Objrs_RitCard As ADODB.Recordset
    Dim M_Objrs_Centra As ADODB.Recordset
    Dim StrKoneksi As String
    
    Dim connIcentra As New ADODB.Connection
    
    
    M_RPTCONN.execute "delete from rptlpddurcall "
    
    '---------------------------------- Cari Yang Di Server 5 ---------------------------------------
    Set connIcentra = New ADODB.Connection
    'Server Icentra 4
    StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
    connIcentra.Open StrKoneksi
    '---------------------------------------------------------------------------------------------------

    cmdsql = "select custid,agent,f_contact_rate from mgm where date(tglcall) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "'  "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where userid between '"
        cmdsql = cmdsql + CStr(Trim(Combo2(0).text)) + "' and '"
        cmdsql = cmdsql + CStr(Trim(Combo2(1).text)) + "' and usertype='1' abd aktif='0') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where spvcode between '"
        cmdsql = cmdsql + CStr(Trim(Combo3(0).text)) + "' and '"
        cmdsql = cmdsql + CStr(Trim(Combo3(1).text)) + "' and usertype='1' and aktif='0') "
    End If
    
    
    Set M_Objrs_RitCard = New ADODB.Recordset
    M_Objrs_RitCard.CursorLocation = adUseClient
    M_Objrs_RitCard.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_Objrs_RitCard.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs_RitCard.RecordCount
        While Not M_Objrs_RitCard.EOF
            ProgressBar1.Value = M_Objrs_RitCard.Bookmark
            'Inputin Data Ke Tabel RptLPDDurCall di Access
            cmdsql = "insert into RptLPDDurCall (custid,agent) values ('"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("custid"))) + "','"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("agent"))) + "')"
            M_RPTCONN.execute cmdsql
            M_Objrs_RitCard.MoveNext
        Wend
        
        'Masukin Deh Durasi Call Yang LPD 1 dan > LPD1 diambil dari Icentra 4
        M_Objrs_RitCard.MoveFirst
        ProgressBar1.Max = M_Objrs_RitCard.RecordCount
        While Not M_Objrs_RitCard.EOF
            ProgressBar1.Value = M_Objrs_RitCard.Bookmark
            
            cmdsql = "select custid,agent,count(dur) as jmlh from report_total_working "
            cmdsql = cmdsql + " where custid='"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("custid"))) + "' and agent='"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_RitCard("agent"))) + "' and "
            cmdsql = cmdsql + " date(calldate) between '"
            cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
            cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "'  "
            cmdsql = cmdsql + " group by custid,agent "
            Set M_Objrs_Centra = New ADODB.Recordset
            M_Objrs_Centra.CursorLocation = adUseClient
            M_Objrs_Centra.Open cmdsql, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_Centra.RecordCount > 0 Then
                cmdsql = " update rptlpddurcall "
                If UCase(Trim(M_Objrs_RitCard("f_contact_rate"))) = "CONTACT LPD" Then
                    cmdsql = cmdsql + " set dur_lpd1='"
                    cmdsql = cmdsql + CStr(M_Objrs_Centra("jmlh")) + "' "
                    cmdsql = cmdsql + " where custid='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("custid")) + "' and agent='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("agent")) + "'"
                    M_RPTCONN.execute cmdsql
                End If
                If UCase(Trim(M_Objrs_RitCard("f_contact_rate"))) = "CONTACT > LPD 1" Then
                    cmdsql = cmdsql + " set dur_lpd1besar='"
                    cmdsql = cmdsql + CStr(M_Objrs_Centra("jmlh")) + "' "
                    cmdsql = cmdsql + " where custid='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("custid")) + "' and agent='"
                    cmdsql = cmdsql + CStr(M_Objrs_RitCard("agent")) + "'"
                    M_RPTCONN.execute cmdsql
                End If
            End If
            Set M_Objrs_Centra = Nothing
            M_Objrs_RitCard.MoveNext
        Wend
        
    End If
    
    Set M_Objrs_RitCard = Nothing
    Set connIcentra = Nothing
End Sub

Private Sub CariCPAApprove()
    Dim M_OBJRS_LPD_LPA As ADODB.Recordset
    Dim SYARAT As String
    Dim cmdsql As String
    Dim Strsql As String
    Dim rsTemp1 As ADODB.Recordset
    Dim rsTemporary As ADODB.Recordset
    
    If TxtCustid.text <> "" Then
        If Len(SYARAT) > 0 Then
            SYARAT = SYARAT + " AND vcustid ='" + TxtCustid.text + "'"
        Else
            SYARAT = " WHERE  vcustid ='" + TxtCustid.text + "'"
        End If
    End If
            

    If Len(SYARAT) > 0 Then
            'SYARAT = SYARAT + " AND dtglinsert  between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
            SYARAT = " WHERE  dpropsal   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
        If Option1(0).Value Then
                SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
                SYARAT = SYARAT + " and userid between '" + Combo2(0).text + "' and '"
                SYARAT = SYARAT + Combo2(1).text + "') "
        End If
        If Option1(1).Value Then
                SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
                SYARAT = SYARAT + " and spvcode between '" + Combo3(0).text + "' and '"
                SYARAT = SYARAT + Combo3(1).text + "') "
        End If
    Else
        'SYARAT = " WHERE  dtglinsert   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
        SYARAT = " WHERE  dpropsal   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
        If Option1(0).Value Then
            SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
            SYARAT = SYARAT + " and userid between '" + Combo2(0).text + "' and '"
            SYARAT = SYARAT + Combo2(1).text + "') "
        End If
        If Option1(1).Value Then
            SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
            SYARAT = SYARAT + " and spvcode between '" + Combo3(0).text + "' and '"
            SYARAT = SYARAT + Combo3(1).text + "') "
        End If
    End If
            
    'RPT.Reset
    M_RPTCONN.execute "delete from tblreportcpa "
    Strsql = "select * from tblreportcpa"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
           
    '@@26Juli2012, Cari Yang Approve digabung Berdasarkan tabel tblsendptp_log_approve
    cmdsql = "select * from "
    
        '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
        cmdsql = cmdsql + "(SELECT  * FROM ( "
        cmdsql = cmdsql + " SELECT * FROM TBLCPA) AS A"
        cmdsql = cmdsql + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID " + SYARAT + ") as cpa_mgm, "
        '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
              
        
        cmdsql = cmdsql + " (select * from tblsendptp_log_approve where date(tgl_proposal)  between '"
        cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "'  and '"
        cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "') as send_ptp_app "
        
    cmdsql = cmdsql + " where cpa_mgm.vcustid=send_ptp_app.custid and "
    cmdsql = cmdsql + " date(cpa_mgm.dpropsal)=date(send_ptp_app.tgl_proposal) "

    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    
    rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = rsTemporary.RecordCount + 1
           
    While Not rsTemporary.EOF
        ProgressBar1.Value = rsTemporary.Bookmark
        DoEvents
        rsTemp1.AddNew
        
        rsTemp1("jenis") = "APPROVED"
        
        rsTemp1("status_ptp") = IIf(IsNull(rsTemporary("status_ptp")), "", rsTemporary("status_ptp"))
        rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
        rsTemp1("dproposal") = IIf(IsNull(rsTemporary("dpropsal")), Null, rsTemporary("dpropsal"))
        rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
        rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
        rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
        rsTemp1("cardno") = IIf(IsNull(rsTemporary("nocard")), "", rsTemporary("nocard"))
        rsTemp1("custname") = IIf(IsNull(rsTemporary("name")), "", rsTemporary("name"))
        rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
        rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
        rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
        rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
        rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
        rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
        rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
        rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
        rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
        rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
        rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
        rsTemp1("custid") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
        rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
        rsTemp1("approve") = IIf(IsNull(rsTemporary("approve_by")), "", rsTemporary("approve_by"))
        rsTemp1("sts_approve") = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
        rsTemp1("payment_after_tenor") = CStr(IIf(IsNull(rsTemporary("payment_after_tenor")), "0", rsTemporary("payment_after_tenor")))
        
        rsTemp1("vjust") = CStr(IIf(IsNull(rsTemporary("vjust")), "", Mid(rsTemporary("vjust"), 1, 250)))
        
        rsTemp1("voccupation") = CStr(IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation")))
        rsTemp1("vreason") = CStr(IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason")))
        
        rsTemp1("chkfaxed") = CStr(IIf(IsNull(rsTemporary("chkfaxed")), "0", rsTemporary("chkfaxed")))
        rsTemp1("chkwentalking") = CStr(IIf(IsNull(rsTemporary("chkwentalking")), "0", rsTemporary("chkwentalking")))
        rsTemp1("chkktp") = CStr(IIf(IsNull(rsTemporary("chkktp")), "0", rsTemporary("chkktp")))
        rsTemp1("chksup") = CStr(IIf(IsNull(rsTemporary("chksup")), "0", rsTemporary("chksup")))
        rsTemp1("chkbillings") = CStr(IIf(IsNull(rsTemporary("chkbillings")), "0", rsTemporary("chkbillings")))
        rsTemp1("chkothers") = CStr(IIf(IsNull(rsTemporary("chkothers")), "0", rsTemporary("chkothers")))
                  
        '@@25072012, Catet f_cek_new yang paid off
        rsTemp1("f_cek_new") = CStr(IIf(Trim(rsTemporary("f_cek_new")) = "PO-", "PO-", ""))
        '@@26Juli2012, Simpan Wo Date nya
        rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, Format(rsTemporary("b_d"), "yyyy-mm-dd"))
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
        cmdsql = "select paydate,payment from tbllunas where custid='"
        cmdsql = cmdsql + CStr(Trim(rsTemporary("custid"))) + "' order by paydate desc limit 1 "

        Set M_OBJRS_LPD_LPA = New ADODB.Recordset
        M_OBJRS_LPD_LPA.CursorLocation = adUseClient
        M_OBJRS_LPD_LPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_OBJRS_LPD_LPA.RecordCount > 0 Then
            rsTemp1("lpd_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("paydate")), Null, Format(M_OBJRS_LPD_LPA("paydate"), "yyyy-mm-dd"))
            rsTemp1("lpa_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("payment")), Null, M_OBJRS_LPD_LPA("payment"))
        Else
            
        End If
        Set M_OBJRS_LPD_LPA = Nothing
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
            rsTemp1.update
            rsTemporary.MoveNext
    Wend
End Sub

'@@011112 Munculkan Data yang akan diajukan ke pak hamanto
Private Sub CariCPA_ToBe_Approve_Hamanto()
    Dim M_OBJRS_LPD_LPA As ADODB.Recordset
    Dim SYARAT As String
    Dim cmdsql As String
    Dim Strsql As String
    Dim rsTemp1 As ADODB.Recordset
    Dim rsTemporary As ADODB.Recordset
    
    
    Strsql = "select * from tblreportcpa"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
           
    cmdsql = "select * from tblsendptp where sts_app_vp='1' and date(tgl_proposal) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and userid between '" + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and spvcode between '" + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    

    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    
    rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = rsTemporary.RecordCount + 1
           
    While Not rsTemporary.EOF
        ProgressBar1.Value = rsTemporary.Bookmark
        DoEvents
        rsTemp1.AddNew
        
        rsTemp1("jenis") = "TO BE APPROVE BY PAK HAMANTO"
        
        rsTemp1("status_ptp") = IIf(IsNull(rsTemporary("jenis_ptp")), "", rsTemporary("jenis_ptp"))
        '@@01-Nov-12, Hanya diambil dari tabel sendptp
        'rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
        rsTemp1("dproposal") = IIf(IsNull(rsTemporary("tgl_proposal")), Null, rsTemporary("tgl_proposal"))
        'rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
        'rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
        rsTemp1("product") = "CARD"
        'rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
        rsTemp1("cardno") = IIf(IsNull(rsTemporary("custid")), "", rsTemporary("custid"))
        rsTemp1("custname") = IIf(IsNull(rsTemporary("vcustname")), "", rsTemporary("vcustname"))
        'rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
        rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
        rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
        rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
        rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
        rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
        rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
        rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
        rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
        rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
        rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
        rsTemp1("custid") = IIf(IsNull(rsTemporary("custid")), "", rsTemporary("custid"))
        rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
        rsTemp1("approve") = IIf(IsNull(rsTemporary("approve_by")), "", rsTemporary("approve_by"))
        'rsTemp1("sts_approve") = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
        rsTemp1("payment_after_tenor") = CStr(IIf(IsNull(rsTemporary("payment_after_tenor")), "0", rsTemporary("payment_after_tenor")))
        
        rsTemp1("vjust") = CStr(IIf(IsNull(rsTemporary("vjust")), "", Mid(rsTemporary("vjust"), 1, 250)))
        
        rsTemp1("voccupation") = CStr(IIf(IsNull(rsTemporary("occupation")), "", rsTemporary("occupation")))
        rsTemp1("vreason") = CStr(IIf(IsNull(rsTemporary("reason")), "", rsTemporary("reason")))
        
        rsTemp1("chkfaxed") = CStr(IIf(IsNull(rsTemporary("chkfaxed")), "0", rsTemporary("chkfaxed")))
        rsTemp1("chkwentalking") = CStr(IIf(IsNull(rsTemporary("chkwentalking")), "0", rsTemporary("chkwentalking")))
        rsTemp1("chkktp") = CStr(IIf(IsNull(rsTemporary("chkktp")), "0", rsTemporary("chkktp")))
        rsTemp1("chksup") = CStr(IIf(IsNull(rsTemporary("chksup")), "0", rsTemporary("chksup")))
        rsTemp1("chkbillings") = CStr(IIf(IsNull(rsTemporary("chkbillings")), "0", rsTemporary("chkbillings")))
        rsTemp1("chkothers") = CStr(IIf(IsNull(rsTemporary("chkothers")), "0", rsTemporary("chkothers")))
                  
        '@@25072012, Catet f_cek_new yang paid off
        'rsTemp1("f_cek_new") = CStr(IIf(Trim(rsTemporary("f_cek_new")) = "PO-", "PO-", ""))
        '@@26Juli2012, Simpan Wo Date nya
        'rsTemp1("wo_date") = IIf(IsNull(rsTemporary("dob")), Null, Format(rsTemporary("dob"), "yyyy-mm-dd"))
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
        cmdsql = "select paydate,payment from tbllunas where custid='"
        cmdsql = cmdsql + CStr(Trim(rsTemporary("custid"))) + "' order by paydate desc limit 1 "

        Set M_OBJRS_LPD_LPA = New ADODB.Recordset
        M_OBJRS_LPD_LPA.CursorLocation = adUseClient
        M_OBJRS_LPD_LPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_OBJRS_LPD_LPA.RecordCount > 0 Then
            rsTemp1("lpd_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("paydate")), Null, Format(M_OBJRS_LPD_LPA("paydate"), "yyyy-mm-dd"))
            rsTemp1("lpa_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("payment")), Null, M_OBJRS_LPD_LPA("payment"))
        Else
            
        End If
        Set M_OBJRS_LPD_LPA = Nothing
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
            rsTemp1.update
            rsTemporary.MoveNext
    Wend
End Sub


Private Sub CariCPAPaidOff()
    Dim M_OBJRS_LPD_LPA As ADODB.Recordset
    Dim SYARAT As String
    Dim cmdsql As String
    Dim Strsql As String
    Dim rsTemp1 As ADODB.Recordset
    Dim rsTemporary As ADODB.Recordset
    
    If TxtCustid.text <> "" Then
        If Len(SYARAT) > 0 Then
            SYARAT = SYARAT + " AND vcustid ='" + TxtCustid.text + "'"
        Else
            SYARAT = " WHERE  vcustid ='" + TxtCustid.text + "'"
        End If
    End If
            

    If Len(SYARAT) > 0 Then
            'SYARAT = SYARAT + " AND dtglinsert  between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
            'SYARAT = " WHERE  dpropsal   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
        If Option1(0).Value Then
                SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
                SYARAT = SYARAT + " and userid between '" + Combo2(0).text + "' and '"
                SYARAT = SYARAT + Combo2(1).text + "') and a.nid in (select max(nid) from tblcpa group by vcustid)  "
        End If
        If Option1(1).Value Then
                SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
                SYARAT = SYARAT + " and spvcode between '" + Combo3(0).text + "' and '"
                SYARAT = SYARAT + Combo3(1).text + "') and a.nid in (select max(nid) from tblcpa group by vcustid) "
        End If
    Else
        'SYARAT = " WHERE  dtglinsert   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
        'SYARAT = " WHERE  dpropsal   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
        If Option1(0).Value Then
            SYARAT = SYARAT + " where b.agent in (select userid from usertbl where usertype='1' "
            SYARAT = SYARAT + " and userid between '" + Combo2(0).text + "' and '"
            SYARAT = SYARAT + Combo2(1).text + "') and a.nid in (select max(nid) from tblcpa group by vcustid) "
        End If
        If Option1(1).Value Then
            SYARAT = SYARAT + " where b.agent in (select userid from usertbl where usertype='1' "
            SYARAT = SYARAT + " and spvcode between '" + Combo3(0).text + "' and '"
            SYARAT = SYARAT + Combo3(1).text + "') and  a.nid in (select max(nid) from tblcpa group by vcustid) "
        End If
    End If
            
    'RPT.Reset
    'M_RPTCONN.Execute "delete from tblreportcpa "
    Strsql = "select * from tblreportcpa"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
           
    '@@26Juli2012, Cari Yang Approve digabung Berdasarkan tabel tblsendptp_log_approve
    cmdsql = "select * from "
    
        '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
        cmdsql = cmdsql + "(SELECT  * FROM ( "
        cmdsql = cmdsql + " SELECT * FROM TBLCPA) AS A"
        'CMDSQL = CMDSQL + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID " + SYARAT + ") as cpa_mgm, "
        '@@22-10-2012 Paid Off jangan diambil dari mgm hst
        cmdsql = cmdsql + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID " + SYARAT + ") as cpa_mgm "
        '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
              
        '@@22-10-2012 Paid Off jangan diambil dari mgm hst
'        CMDSQL = CMDSQL + "(select  custid as custid_hst,tgl,"
'        CMDSQL = CMDSQL + "agent as agent_new from mgm_hst where id in ( "
'        CMDSQL = CMDSQL + "select min(id) from mgm_hst where kodeds like '%PAID OFF%' "
'        CMDSQL = CMDSQL + " and date(tgl) between '"
'        CMDSQL = CMDSQL + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
'        CMDSQL = CMDSQL + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
'        If Option1(0).Value Then
'            CMDSQL = CMDSQL + " and agent in (select userid from usertbl where userid between '"
'            CMDSQL = CMDSQL + Trim(Combo2(0).Text) + "' and '"
'            CMDSQL = CMDSQL + Trim(Combo2(1).Text) + "' and aktif='0' and usertype='1' )  "
'        End If
'        If Option1(1).Value Then
'            CMDSQL = CMDSQL + " and agent in (select userid from usertbl where spvcode between '"
'            CMDSQL = CMDSQL + Trim(Combo3(0).Text) + "' and '"
'            CMDSQL = CMDSQL + Trim(Combo3(1).Text) + "' and aktif='0' and usertype='1' )  "
'        End If
'        CMDSQL = CMDSQL + " group by custid ) "
'        CMDSQL = CMDSQL + ") as paid_off "
        
'    CMDSQL = CMDSQL + " where cpa_mgm.vcustid=paid_off.custid_hst and "
'    CMDSQL = CMDSQL + " cpa_mgm.f_cek_new='PO-'"
     
     cmdsql = cmdsql + " where cpa_mgm.f_cek_new='PO-' and date(tglcall) between '"
     cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
     cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "

    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    
    rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = rsTemporary.RecordCount + 1
           
    While Not rsTemporary.EOF
        ProgressBar1.Value = rsTemporary.Bookmark
        DoEvents
        rsTemp1.AddNew
        
        rsTemp1("jenis") = "PAID OFF"
        
        rsTemp1("status_ptp") = IIf(IsNull(rsTemporary("status_ptp")), "", rsTemporary("status_ptp"))
        rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
        'rsTemp1("dproposal") = IIf(IsNull(rsTemporary("tgl")), Null, rsTemporary("tgl"))
        '@@22102012 dpropsal diambil dari tanggal call aja
        rsTemp1("dproposal") = IIf(IsNull(rsTemporary("tglcall")), Null, rsTemporary("tglcall"))
        rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
        rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
        rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
        rsTemp1("cardno") = IIf(IsNull(rsTemporary("nocard")), "", rsTemporary("nocard"))
        rsTemp1("custname") = IIf(IsNull(rsTemporary("name")), "", rsTemporary("name"))
        rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
        rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
        rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
        rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
        rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
        rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
        rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
        rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
        rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
        rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
        rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
        rsTemp1("custid") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
        rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
        rsTemp1("approve") = IIf(IsNull(rsTemporary("approve_by")), "", rsTemporary("approve_by"))
        rsTemp1("sts_approve") = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
        rsTemp1("payment_after_tenor") = CStr(IIf(IsNull(rsTemporary("payment_after_tenor")), "0", rsTemporary("payment_after_tenor")))
        
        rsTemp1("vjust") = CStr(IIf(IsNull(rsTemporary("vjust")), "", Mid(rsTemporary("vjust"), 1, 250)))
        
        rsTemp1("voccupation") = CStr(IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation")))
        rsTemp1("vreason") = CStr(IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason")))
        
        rsTemp1("chkfaxed") = CStr(IIf(IsNull(rsTemporary("chkfaxed")), "0", rsTemporary("chkfaxed")))
        rsTemp1("chkwentalking") = CStr(IIf(IsNull(rsTemporary("chkwentalking")), "0", rsTemporary("chkwentalking")))
        rsTemp1("chkktp") = CStr(IIf(IsNull(rsTemporary("chkktp")), "0", rsTemporary("chkktp")))
        rsTemp1("chksup") = CStr(IIf(IsNull(rsTemporary("chksup")), "0", rsTemporary("chksup")))
        rsTemp1("chkbillings") = CStr(IIf(IsNull(rsTemporary("chkbillings")), "0", rsTemporary("chkbillings")))
        rsTemp1("chkothers") = CStr(IIf(IsNull(rsTemporary("chkothers")), "0", rsTemporary("chkothers")))
                  
        '@@25072012, Catet f_cek_new yang paid off
        rsTemp1("f_cek_new") = CStr(IIf(Trim(rsTemporary("f_cek_new")) = "PO-", "PO-", ""))
        '@@26Juli2012, Simpan Wo Date nya
        rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, Format(rsTemporary("b_d"), "yyyy-mm-dd"))
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
        cmdsql = "select paydate,payment from tbllunas where custid='"
        cmdsql = cmdsql + CStr(Trim(rsTemporary("custid"))) + "' order by paydate desc limit 1 "

        Set M_OBJRS_LPD_LPA = New ADODB.Recordset
        M_OBJRS_LPD_LPA.CursorLocation = adUseClient
        M_OBJRS_LPD_LPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_OBJRS_LPD_LPA.RecordCount > 0 Then
            rsTemp1("lpd_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("paydate")), Null, Format(M_OBJRS_LPD_LPA("paydate"), "yyyy-mm-dd"))
            rsTemp1("lpa_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("payment")), Null, M_OBJRS_LPD_LPA("payment"))
        Else
            
        End If
        Set M_OBJRS_LPD_LPA = Nothing
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
            rsTemp1.update
            rsTemporary.MoveNext
    Wend
End Sub


Private Sub CariCPARejected()
    Dim M_OBJRS_LPD_LPA As ADODB.Recordset
    Dim SYARAT As String
    Dim cmdsql As String
    Dim Strsql As String
    Dim rsTemp1 As ADODB.Recordset
    Dim rsTemporary As ADODB.Recordset
    
    
    Strsql = "select * from tblreportcpa"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
           
    cmdsql = "select * from tblsendptp_log_reject where date(tgldata) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and userid between '" + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and spvcode between '" + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    
    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    
    rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = rsTemporary.RecordCount + 1
           
    While Not rsTemporary.EOF
        ProgressBar1.Value = rsTemporary.Bookmark
        DoEvents
        rsTemp1.AddNew
        
        rsTemp1("jenis") = "NOT APPROVED"
        
        rsTemp1("status_ptp") = IIf(IsNull(rsTemporary("jenis_ptp")), "", rsTemporary("jenis_ptp"))
        'rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
        rsTemp1("dproposal") = IIf(IsNull(rsTemporary("tgldata")), Null, rsTemporary("tgldata"))
        'rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
        'rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
        'rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
        'rsTemp1("cardno") = IIf(IsNull(rsTemporary("nocard")), "", rsTemporary("nocard"))
        rsTemp1("custname") = IIf(IsNull(rsTemporary("vcustname")), "", rsTemporary("vcustname"))
        'rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
        rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
        rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
        rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
        rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
        rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
        rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
        rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
        rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
        rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
        rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
        rsTemp1("custid") = IIf(IsNull(rsTemporary("custid")), "", rsTemporary("custid"))
        rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("approve_by")), "", rsTemporary("approve_by"))
        'rsTemp1("sts_approve") = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
        rsTemp1("payment_after_tenor") = CStr(IIf(IsNull(rsTemporary("payment_after_tenor")), "0", rsTemporary("payment_after_tenor")))
        
        rsTemp1("vjust") = CStr(IIf(IsNull(rsTemporary("vjust")), "", Mid(rsTemporary("vjust"), 1, 250)))
        
        rsTemp1("voccupation") = CStr(IIf(IsNull(rsTemporary("occupation")), "", rsTemporary("occupation")))
        rsTemp1("vreason") = CStr(IIf(IsNull(rsTemporary("reason")), "", rsTemporary("reason")))
        
        rsTemp1("chkfaxed") = CStr(IIf(IsNull(rsTemporary("chkfaxed")), "0", rsTemporary("chkfaxed")))
        rsTemp1("chkwentalking") = CStr(IIf(IsNull(rsTemporary("chkwentalking")), "0", rsTemporary("chkwentalking")))
        rsTemp1("chkktp") = CStr(IIf(IsNull(rsTemporary("chkktp")), "0", rsTemporary("chkktp")))
        rsTemp1("chksup") = CStr(IIf(IsNull(rsTemporary("chksup")), "0", rsTemporary("chksup")))
        rsTemp1("chkbillings") = CStr(IIf(IsNull(rsTemporary("chkbillings")), "0", rsTemporary("chkbillings")))
        rsTemp1("chkothers") = CStr(IIf(IsNull(rsTemporary("chkothers")), "0", rsTemporary("chkothers")))
                  
        '@@25072012, Catet f_cek_new yang paid off
        'rsTemp1("f_cek_new") = CStr(IIf(Trim(rsTemporary("f_cek_new")) = "PO-", "PO-", ""))
        '@@26Juli2012, Simpan Wo Date nya
        'rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, Format(rsTemporary("b_d"), "yyyy-mm-dd"))
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
        cmdsql = "select paydate,payment from tbllunas where custid='"
        cmdsql = cmdsql + CStr(Trim(rsTemporary("custid"))) + "' order by paydate desc limit 1 "

        Set M_OBJRS_LPD_LPA = New ADODB.Recordset
        M_OBJRS_LPD_LPA.CursorLocation = adUseClient
        M_OBJRS_LPD_LPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_OBJRS_LPD_LPA.RecordCount > 0 Then
            rsTemp1("lpd_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("paydate")), Null, Format(M_OBJRS_LPD_LPA("paydate"), "yyyy-mm-dd"))
            rsTemp1("lpa_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("payment")), Null, M_OBJRS_LPD_LPA("payment"))
        Else
            
        End If
        Set M_OBJRS_LPD_LPA = Nothing
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
            rsTemp1.update
            rsTemporary.MoveNext
    Wend
End Sub


'@@30072012, Buat Narik Keseluruhan CPA
Private Sub TarikSeluruhCPA()
    Dim M_OBJRS_LPD_LPA As ADODB.Recordset
    Dim SYARAT As String
    Dim cmdsql As String
    Dim Strsql As String
    Dim rsTemp1 As ADODB.Recordset
    Dim rsTemporary As ADODB.Recordset
    
'    If TxtCustid.Text <> "" Then
'        If Len(SYARAT) > 0 Then
'            SYARAT = SYARAT + " AND vcustid ='" + TxtCustid.Text + "'"
'        Else
'            SYARAT = " WHERE  vcustid ='" + TxtCustid.Text + "'"
'        End If
'    End If
'
'
'    If Len(SYARAT) > 0 Then
'            'SYARAT = SYARAT + " AND dtglinsert  between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
'            SYARAT = " WHERE  dpropsal   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
'        If Option1(0).Value Then
'                SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
'                SYARAT = SYARAT + " and userid between '" + Combo2(0).Text + "' and '"
'                SYARAT = SYARAT + Combo2(1).Text + "') "
'        End If
'        If Option1(1).Value Then
'                SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
'                SYARAT = SYARAT + " and spvcode between '" + Combo3(0).Text + "' and '"
'                SYARAT = SYARAT + Combo3(1).Text + "') "
'        End If
'    Else
'        'SYARAT = " WHERE  dtglinsert   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
'        SYARAT = " WHERE  dpropsal   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
'        If Option1(0).Value Then
'            SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
'            SYARAT = SYARAT + " and userid between '" + Combo2(0).Text + "' and '"
'            SYARAT = SYARAT + Combo2(1).Text + "') "
'        End If
'        If Option1(1).Value Then
'            SYARAT = SYARAT + " and b.agent in (select userid from usertbl where usertype='1' "
'            SYARAT = SYARAT + " and spvcode between '" + Combo3(0).Text + "' and '"
'            SYARAT = SYARAT + Combo3(1).Text + "') "
'        End If
'    End If
            
    'RPT.Reset
    M_RPTCONN.execute "delete from tblreportcpa "
    Strsql = "select * from tblreportcpa"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
                 
'    '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
'    cmdsql = "SELECT  * FROM ( "
'    cmdsql = cmdsql + " SELECT * FROM TBLCPA where nid in "
'    cmdsql = cmdsql + " (select max(nid) from tblcpa group by vcustid) "
'    cmdsql = cmdsql + ") AS A"
'    cmdsql = cmdsql + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID "
'    '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------

'    '--------------------------- Update Randy Request By:Pak Joko ------------------------
    If Option1(1).Value = True Then
        cmdsql = " SELECT  * FROM ( "
        cmdsql = cmdsql + " SELECT * FROM TBLCPA"
        cmdsql = cmdsql + " where dtglinsert between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'"
        cmdsql = cmdsql + " and '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "') as a"
        cmdsql = cmdsql + " left join mgm b on a.vcustid = b.custid where b.agent in ("
        cmdsql = cmdsql + " select userid from usertbl"
        cmdsql = cmdsql + " where  SPVCODE >='" + Combo3(0).text + "' and SPVCODE <= '" + Combo3(1).text + "')"
        cmdsql = cmdsql + " AND recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        ProgressBar1.Max = rsTemporary.RecordCount + 1
    Else
        cmdsql = " SELECT  * FROM ( "
        cmdsql = cmdsql + " SELECT * FROM TBLCPA"
        cmdsql = cmdsql + " where dtglinsert between '" & Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "'"
        cmdsql = cmdsql + " and '" & Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "') as a"
        cmdsql = cmdsql + " left join mgm b on a.vcustid = b.custid where b.agent in ("
        cmdsql = cmdsql + " select userid from usertbl"
        cmdsql = cmdsql + " where  userid >='" + Combo3(0).text + "' and userid <= '" + Combo3(1).text + "')"
        cmdsql = cmdsql + " AND recsource between  '" + Combo1(0).text + "' and '" + Combo1(1).text + "'"
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        ProgressBar1.Max = rsTemporary.RecordCount + 1
    End If
               
    While Not rsTemporary.EOF
        ProgressBar1.Value = rsTemporary.Bookmark
        DoEvents
        rsTemp1.AddNew
        
        rsTemp1("jenis") = "CPA"
        
        rsTemp1("status_ptp") = IIf(IsNull(rsTemporary("status_ptp")), "", rsTemporary("status_ptp"))
        rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
        rsTemp1("dproposal") = IIf(IsNull(rsTemporary("dpropsal")), Null, rsTemporary("dpropsal"))
        rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
        rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
        rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
        rsTemp1("cardno") = IIf(IsNull(rsTemporary("nocard")), "", rsTemporary("nocard"))
        rsTemp1("custname") = IIf(IsNull(rsTemporary("name")), "", rsTemporary("name"))
        rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
        rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
        rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
        rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
        rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
        rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
        rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
        rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
        rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
        rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
        rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
        rsTemp1("custid") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
        rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
        rsTemp1("approve") = IIf(IsNull(rsTemporary("approve_by")), "", rsTemporary("approve_by"))
        rsTemp1("sts_approve") = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
        rsTemp1("payment_after_tenor") = CStr(IIf(IsNull(rsTemporary("payment_after_tenor")), "0", rsTemporary("payment_after_tenor")))
        
        rsTemp1("vjust") = CStr(IIf(IsNull(rsTemporary("vjust")), "", Mid(rsTemporary("vjust"), 1, 250)))
        
        rsTemp1("voccupation") = CStr(IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation")))
        rsTemp1("vreason") = CStr(IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason")))
        
        rsTemp1("chkfaxed") = CStr(IIf(IsNull(rsTemporary("chkfaxed")), "0", rsTemporary("chkfaxed")))
        rsTemp1("chkwentalking") = CStr(IIf(IsNull(rsTemporary("chkwentalking")), "0", rsTemporary("chkwentalking")))
        rsTemp1("chkktp") = CStr(IIf(IsNull(rsTemporary("chkktp")), "0", rsTemporary("chkktp")))
        rsTemp1("chksup") = CStr(IIf(IsNull(rsTemporary("chksup")), "0", rsTemporary("chksup")))
        rsTemp1("chkbillings") = CStr(IIf(IsNull(rsTemporary("chkbillings")), "0", rsTemporary("chkbillings")))
        rsTemp1("chkothers") = CStr(IIf(IsNull(rsTemporary("chkothers")), "0", rsTemporary("chkothers")))
                  
        '@@25072012, Catet f_cek_new yang paid off
        rsTemp1("f_cek_new") = CStr(IIf(Trim(rsTemporary("f_cek_new")) = "PO-", "PO-", ""))
        '@@26Juli2012, Simpan Wo Date nya
        rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, Format(rsTemporary("b_d"), "yyyy-mm-dd"))
        
'        '-------------------------------- Cari Total Payment ----------------------------------------
'        CMDSQL = "select sum(payment) as lunas from tbllunas where custid='"
'        CMDSQL = CMDSQL + CStr(Trim(rsTemporary("custid"))) + "'   "
'
'        Set M_OBJRS_LPD_LPA = New ADODB.Recordset
'        M_OBJRS_LPD_LPA.CursorLocation = adUseClient
'        M_OBJRS_LPD_LPA.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If M_OBJRS_LPD_LPA.RecordCount > 0 Then
'            rsTemp1("total_lunas") = IIf(IsNull(M_OBJRS_LPD_LPA("lunas")), Null, M_OBJRS_LPD_LPA("lunas"))
'        Else
'
'        End If
'        Set M_OBJRS_LPD_LPA = Nothing
'
'        '-------------------------------- Cari Total Payment ----------------------------------------
            rsTemp1.update
            rsTemporary.MoveNext
    Wend
    
    If b_excel Then
        'Call UpdateAllPaymentCPA
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPTPJatuhTempoNew_REGULER.rpt"
        RPT.RetrieveDataFiles
        RPT.Destination = crptToFile
        RPT.PrintFileType = crptExcel50
        RPT.action = 1
    End If

End Sub

Private Sub UpdateAllPaymentCPA()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmdsql = "select mgm.custid,sum(tbllunas.payment) as jumlah from tbllunas,mgm "
    cmdsql = cmdsql + " where mgm.custid=tbllunas.custid group by mgm.custid "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        ProgressBar1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            ProgressBar1.Value = M_Objrs.Bookmark
            cmdsql = "update tblreportcpa set total_lunas='"
            cmdsql = cmdsql + CStr(M_Objrs("jumlah")) + "' where custid='"
            cmdsql = cmdsql + CStr(Trim(M_Objrs("custid"))) + "'"
            M_RPTCONN.execute cmdsql
M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub
'@@ 19-04-2013, detail payment interval --> By : Budi <--
Private Sub RptDetailPayment_Interval_permonth()


Dim RsGetAllData As ADODB.Recordset
Dim RsAccessData As ADODB.Recordset
Dim cmdsql As String


    On Error GoTo adderr

  

    cmdsql = " select a.*, mgm.name, mgm.acc_type from mgm,( SELECT m.custid, m.tahun, COALESCE(m.""1"", 0) AS ""Jan"","
    cmdsql = cmdsql + " COALESCE(m.""2"", 0) AS ""Feb"", COALESCE(m.""3"", 0) AS ""Mar"", COALESCE(m.""4"", 0) AS ""Apr"", COALESCE(m.""5"", 0) AS ""Mei"","
    cmdsql = cmdsql + " COALESCE(m.""6"", 0) AS ""Jun"", COALESCE(m.""7"", 0) AS ""Jul"", COALESCE(m.""8"", 0) AS ""Aug"", COALESCE(m.""9"", 0) AS ""Sep"","
    cmdsql = cmdsql + " COALESCE(m.""10"", 0) AS ""Okt"", COALESCE(m.""11"", 0) AS ""Nop"", COALESCE(m.""12"", 0) AS ""Des""  "
    cmdsql = cmdsql + " FROM crosstab('select date_part(''year'',paydate)as tahun, custid, date_part(''month'',paydate) as bulan, "
    cmdsql = cmdsql + " sum(payment) as payment from tbllunas where paydate between "
    cmdsql = cmdsql + " ''" + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "'' and "
    cmdsql = cmdsql + " ''" + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "'' "
    
    If Option1(0).Value Then
        cmdsql = cmdsql + "  and custid  in ( select custid from tbllunas where agent in "
        cmdsql = cmdsql + " (select userid from usertbl where usertype=''1'' "
        cmdsql = cmdsql + " and userid between ''" + Combo2(0).text + "'' and "
        cmdsql = cmdsql + " ''" + Combo2(1).text + "'') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + "  and custid  in ( select custid from tbllunas where agent in "
        cmdsql = cmdsql + "  (select userid from usertbl where usertype=''1'' "
        cmdsql = cmdsql + " and spvcode between ''" + Combo3(0).text + "'' and "
        cmdsql = cmdsql + " ''" + Combo3(1).text + "'') "
    End If
    cmdsql = cmdsql + " ) group by tahun,custid,bulan order by custid,tahun, bulan'::text, 'select m from generate_series(1,12) m'::text) m(""tahun"" text, ""custid"" text, ""1"" numeric, ""2"" numeric, ""3"" numeric, ""4"" numeric, ""5"" numeric, ""6"" numeric, ""7"" numeric, ""8"""
    cmdsql = cmdsql + " numeric, ""9"" numeric, ""10"" numeric, ""11"" numeric, ""12"" numeric)) a where a.custid=mgm.custid"
    Set RsGetAllData = New ADODB.Recordset
    RsGetAllData.CursorLocation = adUseClient
    RsGetAllData.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    Set RsAccessData = New ADODB.Recordset
    RsAccessData.CursorLocation = adUseClient
    M_RPTCONN.execute "delete from rptdetailpayment"
    RsAccessData.Open "select * from rptdetailpayment", M_RPTCONN, adOpenDynamic, adLockOptimistic
    If RsGetAllData.RecordCount > 0 Then
        ProgressBar1.Max = RsGetAllData.RecordCount
    Else
        MsgBox "Data yang dicari tidak ditemukan!!!", vbOKOnly + vbInformation, "INFO"
        Exit Sub
    End If
    While Not RsGetAllData.EOF
        ProgressBar1.Value = RsGetAllData.Bookmark
        RsAccessData.AddNew
        RsAccessData!CustId = RsGetAllData!CustId
        RsAccessData!Name = RsGetAllData!Name
        RsAccessData!PRODUCT = RsGetAllData!acc_type
        RsAccessData!Tahun = RsGetAllData!Tahun
        RsAccessData!Jan = RsGetAllData!Jan
        RsAccessData!feb = RsGetAllData!feb
        RsAccessData!Mar = RsGetAllData!Mar
        RsAccessData!Apr = RsGetAllData!Apr
        RsAccessData!Mei = RsGetAllData!Mei
        RsAccessData!Jun = RsGetAllData!Jun
        RsAccessData!Jul = RsGetAllData!Jul
        RsAccessData!Aug = RsGetAllData!Aug
        RsAccessData!Sep = RsGetAllData!Sep
        RsAccessData!Okt = RsGetAllData!Okt
        RsAccessData!Nop = RsGetAllData!Nop
        RsAccessData!des = RsGetAllData!des
        RsAccessData!CODE = IIf(RsGetAllData!Jan = 0, "O", "X") & IIf(RsGetAllData!feb = 0, "O", "X") & _
                            IIf(RsGetAllData!Mar = 0, "O", "X") & IIf(RsGetAllData!Apr = 0, "O", "X") & _
                            IIf(RsGetAllData!Mei = 0, "O", "X") & IIf(RsGetAllData!Jun = 0, "O", "X") & _
                            IIf(RsGetAllData!Jul = 0, "O", "X") & IIf(RsGetAllData!Aug = 0, "O", "X") & _
                            IIf(RsGetAllData!Sep = 0, "O", "X") & IIf(RsGetAllData!Okt = 0, "O", "X") & _
                            IIf(RsGetAllData!Nop = 0, "O", "X") & IIf(RsGetAllData!des = 0, "O", "X")
                            
                            
        RsAccessData.update
    RsGetAllData.MoveNext
    
    Wend
    
    

    Dim lc, NxtLine, K
        Screen.MousePointer = vbHourglass
        Set ExlObj = CreateObject("excel.application")      ' Initialize the excel object
        ExlObj.Workbooks.ADD                                ' Add an excel workbook
        ' Get the required data from the database
        RsAccessData.Requery
        If Not RsAccessData.EOF Then
          ExlObj.Visible = True         ' Show the excel sheet
          With ExlObj.ActiveSheet
            
            ' Print the heading and columns first
                    
            .Cells(1, 3).Value = "Report Detail Payment Interval Permonth"
            .Cells(1, 3).Font.Name = "Verdana"
            .Cells(1, 3).Font.Bold = True:
            .Cells(2, 3).Value = "Periode : " & Format(TDBDate1(0).Value, "dd-mm-yyyy") & "s/d" & Format(TDBDate1(1).Value, "dd-mm-yyyy")
            .Cells(2, 3).Font.Name = "Verdana"
            .Cells(2, 3).Font.Bold = True:
            .Cells(4, 1).Value = "CustId":    .Cells(4, 2).Value = "Name"
            .Cells(4, 3).Value = "Product"
            .Cells(4, 4).Value = "Code":      .Cells(4, 5).Value = "Tahun"
            .Cells(4, 6).Value = "Jan":       .Cells(4, 7).Value = "Feb"
            .Cells(4, 8).Value = "Mar":       .Cells(4, 9).Value = "Apr"
            .Cells(4, 10).Value = "Mei":       .Cells(4, 11).Value = "Jun"
            .Cells(4, 12).Value = "Jul":       .Cells(4, 13).Value = "Aug"
            .Cells(4, 14).Value = "Sep":       .Cells(4, 15).Value = "Okt"
            .Cells(4, 16).Value = "Nop":       .Cells(4, 17).Value = "Des"
            
          End With
        End If
        For K = 1 To RsGetAllData.fields.Count
            ' Column headings are set to bold and white.
            ExlObj.ActiveSheet.Cells(4, K).Font.Bold = True
            ExlObj.ActiveSheet.Cells(4, K).Font.Color = vbBlue
        Next
        Set K = Nothing
        NxtLine = 5
        
        ' Now we will export data into the sheetr
        'RsAccessData.Requery
        If RsAccessData.RecordCount > 0 Then
        ProgressBar1.Max = RsAccessData.RecordCount
        End If
        Do Until RsAccessData.EOF
            ProgressBar1.Value = RsAccessData.Bookmark
            For lc = 0 To RsAccessData.fields.Count - 1
                ExlObj.ActiveSheet.Cells(NxtLine, lc + 1).Value = RsAccessData.fields(lc)
                If RsAccessData.fields.Item(lc).Name <> "DATE" Then
                   ExlObj.ActiveSheet.Cells(NxtLine, lc + 1).Value = RsAccessData.fields(lc)
                Else
                   ExlObj.ActiveSheet.Cells(NxtLine, lc + 1).Value = Format(RsAccessData.fields(lc), "dd/mm/yy")
                End If
                
            Next
            RsAccessData.MoveNext
            NxtLine = NxtLine + 1
        Loop
        
        ' Once the data has been exported, we will format the sheet _
          by using the AutoFormat function.
        ExlObj.ActiveCell.Worksheet.Cells(NxtLine, lc + 1).AutoFormat xlRangeAutoFormatList3, 0, , 3, 1, True, True
        'ExlObj.ActiveCell.Worksheet.Cells.AutoFormat  '<- Click the space key after _
                                                         .AutoFormat to see its _
                                                         parameter types.

        Set RsAccessData = Nothing
        Set RsGetAllData = Nothing
        Set ExlObj = Nothing
        Screen.MousePointer = vbDefault
        MsgBox "Retreive Done"
    Exit Sub
adderr:
    MsgBox err.Description

End Sub

'@@26022013 Ini tambahan Report Request PTP
Private Sub RptRequestPTP()
    Dim M_OBJRS_LPD_LPA As ADODB.Recordset
    Dim SYARAT As String
    Dim cmdsql As String
    Dim Strsql As String
    Dim rsTemp1 As ADODB.Recordset
    Dim rsTemporary As ADODB.Recordset
    Dim Tanggal As String
      
    Dim M_WHERE As String
    Dim jenis As String
    
    On Error GoTo SALAH
    
    Strsql = "delete from tblreportcpa "
    M_RPTCONN.execute Strsql
    
    Strsql = "select * from tblreportcpa"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open Strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
           
    cmdsql = "select * from tblsendptp as ts,mgm where date(ts.tgldata) between '"
    cmdsql = cmdsql + Format(TDBDate1(0).Value, "yyyy-mm-dd") + "' and '"
    cmdsql = cmdsql + Format(TDBDate1(1).Value, "yyyy-mm-dd") + "' "
    If Option1(0).Value Then
        cmdsql = cmdsql + " and ts.agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and userid between '" + Combo2(0).text + "' and '"
        cmdsql = cmdsql + Combo2(1).text + "') "
    End If
    If Option1(1).Value Then
        cmdsql = cmdsql + " and ts.agent in (select userid from usertbl where usertype='1' "
        cmdsql = cmdsql + " and spvcode between '" + Combo3(0).text + "' and '"
        cmdsql = cmdsql + Combo3(1).text + "') "
    End If
    
    M_WHERE = " and mgm.custid=ts.custid "
    
    
    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    
    rsTemporary.Open cmdsql + M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ProgressBar1.Max = rsTemporary.RecordCount + 1
               
    While Not rsTemporary.EOF
        ProgressBar1.Value = rsTemporary.Bookmark
        DoEvents
        rsTemp1.AddNew
        
        If IsNull(rsTemporary("approve_by")) = True Then
            jenis = "NOT APPROVED"
            rsTemp1("dproposal") = IIf(IsNull(rsTemporary("tgldata")), Null, rsTemporary("tgldata"))
        Else
            jenis = rsTemporary("approve_by")
            rsTemp1("dproposal") = IIf(IsNull(rsTemporary("tglproposal")), Null, rsTemporary("tglproposal"))
        End If
        
        'rsTemp1("jenis") = "NOT APPROVED"
        rsTemp1("jenis") = jenis
        
        rsTemp1("status_ptp") = IIf(IsNull(rsTemporary("jenis_ptp")), "", rsTemporary("jenis_ptp"))
        'rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
        
        'rsTemp1("dproposal") = IIf(IsNull(rsTemporary("tgldata")), Null, rsTemporary("tgldata"))
        
        'rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
        rsTemp1("product") = IIf(IsNull(rsTemporary("acc_type")), "", rsTemporary("acc_type"))
        'rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
        'rsTemp1("cardno") = IIf(IsNull(rsTemporary("nocard")), "", rsTemporary("nocard"))
        rsTemp1("custname") = IIf(IsNull(rsTemporary("vcustname")), "", rsTemporary("vcustname"))
        'rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
        rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
        rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
        rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
        rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
        rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
        rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
        rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
        rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
        rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
        rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
        rsTemp1("custid") = IIf(IsNull(rsTemporary("custid")), "", rsTemporary("custid"))
        rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
        'rsTemp1("approve") = IIf(IsNull(rsTemporary("approve_by")), "", rsTemporary("approve_by"))
        'rsTemp1("sts_approve") = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
        rsTemp1("payment_after_tenor") = CStr(IIf(IsNull(rsTemporary("payment_after_tenor")), "0", rsTemporary("payment_after_tenor")))
        
        rsTemp1("vjust") = CStr(IIf(IsNull(rsTemporary("vjust")), "", Mid(rsTemporary("vjust"), 1, 250)))
        
        rsTemp1("voccupation") = CStr(IIf(IsNull(rsTemporary("occupation")), "", rsTemporary("occupation")))
        rsTemp1("vreason") = CStr(IIf(IsNull(rsTemporary("reason")), "", rsTemporary("reason")))
        
        rsTemp1("chkfaxed") = CStr(IIf(IsNull(rsTemporary("chkfaxed")), "0", rsTemporary("chkfaxed")))
        rsTemp1("chkwentalking") = CStr(IIf(IsNull(rsTemporary("chkwentalking")), "0", rsTemporary("chkwentalking")))
        rsTemp1("chkktp") = CStr(IIf(IsNull(rsTemporary("chkktp")), "0", rsTemporary("chkktp")))
        rsTemp1("chksup") = CStr(IIf(IsNull(rsTemporary("chksup")), "0", rsTemporary("chksup")))
        rsTemp1("chkbillings") = CStr(IIf(IsNull(rsTemporary("chkbillings")), "0", rsTemporary("chkbillings")))
        rsTemp1("chkothers") = CStr(IIf(IsNull(rsTemporary("chkothers")), "0", rsTemporary("chkothers")))
                  
        '@@25072012, Catet f_cek_new yang paid off
        'rsTemp1("f_cek_new") = CStr(IIf(Trim(rsTemporary("f_cek_new")) = "PO-", "PO-", ""))
        '@@26Juli2012, Simpan Wo Date nya
        'rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, Format(rsTemporary("b_d"), "yyyy-mm-dd"))
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
        cmdsql = "select paydate,payment from tbllunas where custid='"
        cmdsql = cmdsql + CStr(Trim(rsTemporary("custid"))) + "' order by paydate desc limit 1 "

        Set M_OBJRS_LPD_LPA = New ADODB.Recordset
        M_OBJRS_LPD_LPA.CursorLocation = adUseClient
        M_OBJRS_LPD_LPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_OBJRS_LPD_LPA.RecordCount > 0 Then
            rsTemp1("lpd_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("paydate")), Null, Format(M_OBJRS_LPD_LPA("paydate"), "yyyy-mm-dd"))
            rsTemp1("lpa_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("payment")), Null, M_OBJRS_LPD_LPA("payment"))
        Else
            
        End If
        Set M_OBJRS_LPD_LPA = Nothing
        
        '-------------------------------- Cari LPD dan LPA ----------------------------------------
            rsTemp1.update
            rsTemporary.MoveNext
    Wend
    
    If b_excel Then
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Rptlist.rpt"
        RPT.RetrieveDataFiles
        RPT.Destination = crptToFile
        RPT.PrintFileType = crptExcel50
        RPT.action = 1
    End If
    
    Exit Sub
SALAH:
    MsgBox "Maaf, ada kesalahan! " & err.Description, vbOKOnly + vbExclamation, "Informasi"
    
End Sub

Private Sub Export_Excel(M_Objrs As ADODB.Recordset)
    On Error GoTo SALAH
    Dim cmdsql As String
    Dim listItem As listItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i               As Integer
    Dim m_msgbox        As String
    
    i = 1

form_save:
    CD_save.ShowSave
    Txtpath.text = CD_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo form_save        '-> maka goto form_save
        End If
    End If

    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    On Error GoTo SALAH
    
    Dim x, Y    As Integer
    If M_Objrs.state = 1 Then
        x = 0
        Y = M_Objrs.fields().Count - 1
        Do Until x > Y
            DoEvents
            objSheet.Cells(1, i).Value = CStr(M_Objrs.fields(x).Name)
            i = i + 1
            x = x + 1
        Loop
    End If
    
    objSheet.Range("A2").CopyFromRecordset M_Objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs Txtpath.text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_Objrs = Nothing
    
    Exit Sub
SALAH:
    MsgBox err.Description
    Set M_Objrs = Nothing
    Exit Sub
End Sub
