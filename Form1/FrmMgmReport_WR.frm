VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmMgmReport_AWARNESS 
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11715
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport RPT 
      Left            =   6600
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Data yang sdh di tarik"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   36
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Data yang belum tarik"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pil"
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   34
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Rit"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   33
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Awarness"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CmbCek 
      Height          =   315
      Left            =   9240
      TabIndex        =   24
      Top             =   3210
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   360
      Index           =   0
      Left            =   9165
      TabIndex        =   12
      Top             =   3630
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   10320
      TabIndex        =   13
      Top             =   3600
      Width           =   1125
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   9300
      TabIndex        =   7
      Top             =   1695
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   6570
      TabIndex        =   6
      Top             =   1695
      Width           =   2505
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose One..."
      Height          =   1035
      Left            =   5760
      TabIndex        =   14
      Top             =   570
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
         Top             =   195
         Width           =   2085
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   1
         Left            =   3630
         TabIndex        =   5
         Top             =   540
         Width           =   2130
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   4
         Top             =   540
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Agent        :"
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Supervisor :"
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   555
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   2
         Left            =   3375
         TabIndex        =   18
         Top             =   225
         Width           =   270
      End
      Begin VB.Label Label1 
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
      Height          =   3405
      Left            =   -30
      TabIndex        =   16
      Top             =   -15
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   6006
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
      Left            =   5760
      TabIndex        =   15
      Top             =   4005
      Visible         =   0   'False
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   1
      Left            =   9360
      TabIndex        =   10
      Top             =   2040
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport_WR.frx":0000
      Caption         =   "FrmMgmReport_WR.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport_WR.frx":0184
      Keys            =   "FrmMgmReport_WR.frx":01A2
      Spin            =   "FrmMgmReport_WR.frx":0200
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
      Top             =   2040
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport_WR.frx":0228
      Caption         =   "FrmMgmReport_WR.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport_WR.frx":03AC
      Keys            =   "FrmMgmReport_WR.frx":03CA
      Spin            =   "FrmMgmReport_WR.frx":0428
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
      Left            =   8160
      TabIndex        =   9
      Top             =   2025
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   529
      Caption         =   "FrmMgmReport_WR.frx":0450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport_WR.frx":04BC
      Spin            =   "FrmMgmReport_WR.frx":050C
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
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   2010382337
      Value           =   2.12482692446619E-314
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   1
      Left            =   10680
      TabIndex        =   11
      Top             =   2010
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReport_WR.frx":0534
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport_WR.frx":05A0
      Spin            =   "FrmMgmReport_WR.frx":05F0
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
      Top             =   2640
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport_WR.frx":0618
      Caption         =   "FrmMgmReport_WR.frx":0730
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport_WR.frx":079C
      Keys            =   "FrmMgmReport_WR.frx":07BA
      Spin            =   "FrmMgmReport_WR.frx":0818
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
      Left            =   9360
      TabIndex        =   27
      Top             =   2640
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   556
      Calendar        =   "FrmMgmReport_WR.frx":0840
      Caption         =   "FrmMgmReport_WR.frx":0958
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReport_WR.frx":09C4
      Keys            =   "FrmMgmReport_WR.frx":09E2
      Spin            =   "FrmMgmReport_WR.frx":0A40
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
      Left            =   8160
      TabIndex        =   28
      Top             =   2640
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   529
      Caption         =   "FrmMgmReport_WR.frx":0A68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport_WR.frx":0AD4
      Spin            =   "FrmMgmReport_WR.frx":0B24
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
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   2010382337
      Value           =   2.12482692446619E-314
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   3
      Left            =   10680
      TabIndex        =   29
      Top             =   2640
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReport_WR.frx":0B4C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReport_WR.frx":0BB8
      Spin            =   "FrmMgmReport_WR.frx":0C08
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
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   2010382337
      Value           =   2.12482692446619E-314
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   300
      Index           =   7
      Left            =   9120
      TabIndex        =   31
      Top             =   2640
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Comparator Date :"
      Height          =   420
      Index           =   3
      Left            =   5520
      TabIndex        =   30
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Status Cek :"
      Height          =   315
      Left            =   7875
      TabIndex        =   25
      Top             =   3225
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "to"
      Height          =   255
      Index           =   1
      Left            =   8970
      TabIndex        =   23
      Top             =   1710
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From Batch :"
      Height          =   300
      Index           =   0
      Left            =   5520
      TabIndex        =   22
      Top             =   1725
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   300
      Index           =   5
      Left            =   5640
      TabIndex        =   21
      Top             =   2055
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   300
      Index           =   4
      Left            =   9090
      TabIndex        =   20
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   120
      Width           =   5745
   End
End
Attribute VB_Name = "FrmMgmReport_AWARNESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_OBJRS As New ADODB.Recordset
Dim JUMLAH As String
Dim JUMLAHVOL As String
Dim batch As String
Dim CMDSQL As String
Dim CMDSQL1 As String
Dim STATUS As String
Dim LAmount As String
Dim LAgent As String
Dim LAgent1 As String
Dim Last As String
Dim Jml As String
Dim Lf_cek As String
Dim Lvol As String
Dim TOTPTP As String
Dim dtcek As Boolean
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
CMDSQL = "SELECT mgm_hst.agent, sum(mgm.amountwo)as jmlAmount from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and "
CMDSQL = CMDSQL + "custid in(select custid from mgm where agent between '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "') "
CMDSQL = CMDSQL + "group by custid,agent) as a on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax "
CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where  recsource BETWEEN "
CMDSQL = CMDSQL + " '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent order by mgm_hst.agent"
Else
If Option1(1).Value Then
CMDSQL = "SELECT mgm_hst.agent,  sum(mgm.amountwo)as jmlAmount from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where  tgl BETWEEN"
CMDSQL = CMDSQL + "'" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and "
CMDSQL = CMDSQL + "custid in(select custid from mgm where agent in (select userid from usertbl where "
CMDSQL = CMDSQL + "SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')) group by custid,agent) as a "
CMDSQL = CMDSQL + "on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  inner join "
CMDSQL = CMDSQL + "mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
CMDSQL = CMDSQL + "RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent "
End If
End If
DoEvents
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 2
While Not m_hst.EOF
ProgressBar1.Value = m_hst.Bookmark
LAgent1 = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
LAmount = Trim(CStr(IIf(IsNull(m_hst!jmlamount), 0, m_hst!jmlamount)))
CMDSQL1 = " Update TrackingRptPerPrgBatch set VOLUTILIZED= '" + LAmount + "' where AOC='" + LAgent1 + "'"
M_RPTCONN.Execute CMDSQL1
m_hst.MoveNext
Wend
Set m_hst = Nothing
CMDSQL = Empty
CMDSQL1 = Empty
LAgent1 = Empty

Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description & "eror di volume utilzed"
    End If
    Resume Next
End Sub

Private Sub VolumeUtilized_TARIK()
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
CMDSQL = "SELECT mgm_hst.agent, sum(mgmTARIK.amountwo)as jmlAmount from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and "
CMDSQL = CMDSQL + "custid in(select custid from mgmTARIK where agent between '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "') "
CMDSQL = CMDSQL + "group by custid,agent) as a on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax "
CMDSQL = CMDSQL + "inner join mgmTARIK on mgmTARIK.custid = a.custid and a.agent=mgmTARIK.agent where  recsource BETWEEN "
CMDSQL = CMDSQL + " '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent order by mgm_hst.agent"
Else
If Option1(1).Value Then
CMDSQL = "SELECT mgm_hst.agent,  sum(mgmTARIK.amountwo)as jmlAmount from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where  tgl BETWEEN"
CMDSQL = CMDSQL + "'" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and "
CMDSQL = CMDSQL + "custid in(select custid from mgmTARIK where agent in (select userid from usertbl where "
CMDSQL = CMDSQL + "SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')) group by custid,agent) as a "
CMDSQL = CMDSQL + "on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  inner join "
CMDSQL = CMDSQL + "mgmTARIK on mgmTARIK.custid = a.custid and a.agent=mgmTARIK.agent where "
CMDSQL = CMDSQL + "RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent order by mgm_hst.agent "
End If
End If

m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 2
While Not m_hst.EOF
ProgressBar1.Value = m_hst.Bookmark
LAgent1 = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
LAmount = Trim(CStr(IIf(IsNull(m_hst!jmlamount), 0, m_hst!jmlamount)))
CMDSQL1 = " Update TrackingRptPerPrgBatch set VOLUTILIZED= '" + LAmount + "' where AOC='" + LAgent1 + "'"
M_RPTCONN.Execute CMDSQL1
m_hst.MoveNext
Wend
Set m_hst = Nothing
CMDSQL = Empty
CMDSQL1 = Empty
LAgent1 = Empty

Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next
End Sub

Private Sub ReportPTPNego()
Dim Rsptp As ADODB.Recordset
Dim m_msgbox As Variant

On Error GoTo eddder:
Set Rsptp = New ADODB.Recordset
Rsptp.CursorLocation = adUseClient
If Option1(0).Value Then
CMDSQL = "select agent,f_cek,count(agent) as JML,sum(promisepay) as VOL from reportPTP  where"
CMDSQL = CMDSQL + " agent in (select userid from usertbl where userid >='" + Combo2(0).Text + "' and userid<='" + Combo2(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + "  group by agent, f_cek "

Else
If Option1(1).Value Then
CMDSQL = "select agent,f_cek, count(agent) as JML,sum(promisepay) as VOL from reportPTP  where"
CMDSQL = CMDSQL + " agent in (select userid from usertbl where spvcode >='" + Combo3(0).Text + "' and SPVCODE<='" + Combo3(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + "  group by agent,f_cek "
End If
End If

Rsptp.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not Rsptp.EOF
LAgent = Trim(IIf(IsNull(Rsptp!agent), "", Rsptp!agent))
Jml = Trim(IIf(IsNull(Rsptp!Jml), 0, Rsptp!Jml))
Lf_cek = Trim(IIf(IsNull(Rsptp!F_CEK), "", Rsptp!F_CEK))
Lvol = Trim(IIf(IsNull(Rsptp!vol), 0, Rsptp!vol))
If Left(Lf_cek, 3) = "PTP" Then
Jml = Val(Jml) + 1
TOTPTP = Val(TOTPTP) + Val(Lvol)
Else
Jml = 0
TOTPTP = 0
End If
M_RPTCONN.Execute "UPDATE TrackingRptPerPrgBatch set PTP_BARU =" + Jml + ",VolPTP_Baru=" + TOTPTP + "  where AOC = '" + LAgent + "'"
Rsptp.MoveNext
Wend
Set Rsptp = Nothing
CMDSQL = Empty
LAgent = Empty
Jml = Empty
Lf_cek = Empty
TOTPTP = Empty
Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next

End Sub

Private Sub ReportPTPNego_TARIK()
Dim Rsptp As ADODB.Recordset
Dim m_msgbox As Variant

On Error GoTo eddder:
Set Rsptp = New ADODB.Recordset
Rsptp.CursorLocation = adUseClient
If Option1(0).Value Then
CMDSQL = "select agent,f_cek,count(agent) as JML,sum(promisepay) as VOL from reportPTPTARIK  where "
CMDSQL = CMDSQL + " agent in (select userid from usertbl where userid >='" + Combo2(0).Text + "' and userid<='" + Combo2(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + "  group by agent, f_cek "

Else
If Option1(1).Value Then
CMDSQL = "select agent,f_cek, count(agent) as JML,sum(promisepay) as VOL from reportPTPTARIK  where "
CMDSQL = CMDSQL + " agent in (select userid from usertbl where spvcode >='" + Combo3(0).Text + "' and SPVCODE<='" + Combo3(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + "  group by agent,f_cek "
End If
End If

Rsptp.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not Rsptp.EOF
LAgent = Trim(IIf(IsNull(Rsptp!agent), "", Rsptp!agent))
Jml = Trim(IIf(IsNull(Rsptp!Jml), 0, Rsptp!Jml))
Lf_cek = Trim(IIf(IsNull(Rsptp!F_CEK), "", Rsptp!F_CEK))
Lvol = Trim(IIf(IsNull(Rsptp!vol), 0, Rsptp!vol))
If Lf_cek = "PTP" Then
Jml = Jml
Lvol = Lvol
Else
Jml = 0
Lvol = 0
End If
M_RPTCONN.Execute "UPDATE TrackingRptPerPrgBatch set PTP_BARU =" + Jml + ",VolPTP_Baru=" + Lvol + "  where AOC = '" + LAgent + "'"
Rsptp.MoveNext
Wend
Set Rsptp = Nothing
CMDSQL = Empty
LAgent = Empty
Jml = Empty
Lf_cek = Empty

Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
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
CMDSQL = "select agent,f_cek,count(agent) as JML,sum(promisepay) as VOL from reportPTP  where "
CMDSQL = CMDSQL + " agent in (select userid from usertbl where userid >='" + Combo2(0).Text + "' and userid<='" + Combo2(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "' "
CMDSQL = CMDSQL + "  group by agent, f_cek "

Else
If Option1(1).Value Then
CMDSQL = "select agent,f_cek, count(agent) as JML,sum(promisepay) as VOL from reportPTP  where "
CMDSQL = CMDSQL + " agent in (select userid from usertbl where spvcode >='" + Combo3(0).Text + "' and SPVCODE<='" + Combo3(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "' "
CMDSQL = CMDSQL + "  group by agent,f_cek "
End If
End If

Rsptp.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not Rsptp.EOF
LAgent = Trim(IIf(IsNull(Rsptp!agent), "", Rsptp!agent))
Jml = Trim(IIf(IsNull(Rsptp!Jml), 0, Rsptp!Jml))
Lf_cek = Trim(IIf(IsNull(Rsptp!F_CEK), "", Rsptp!F_CEK))
Lvol = Trim(IIf(IsNull(Rsptp!vol), 0, Rsptp!vol))
If Lf_cek = "PTP" Then
Jml = Jml
Else
Jml = 0
End If
M_RPTCONN.Execute "UPDATE TrackingRptPerPrgBatch set PTP_BARU_LM =" + Jml + ",VolPTP_Baru_LM=" + Lvol + "  where AOC = '" + LAgent + "'"
Rsptp.MoveNext
Wend
Set Rsptp = Nothing
CMDSQL = Empty
LAgent = Empty
Jml = Empty
Lf_cek = Empty

Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
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

If Option1(0).Value Then

CMDSQL = "SELECT AGENT, F_CEK, COUNT(F_CEK) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,F_CEK, agent,ttlptp from mgm"
CMDSQL = CMDSQL + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
'CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
CMDSQL = CMDSQL + " AND (left(F_CEK,2) IN ('NK','MV','WN','SP','NA','RP','BP','OP','ST')"
CMDSQL = CMDSQL + " or left(F_CEK,3) in ('NBP','PTP','POP','PRE'))"
CMDSQL = CMDSQL + " AND custid in (Select distinct custid from mgm_hst"
CMDSQL = CMDSQL + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT,f_cek"

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

CMDSQL = "SELECT AGENT, F_CEK, COUNT(F_CEK) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,F_CEK, agent,ttlptp from mgm"
CMDSQL = CMDSQL + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
'CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
'CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
CMDSQL = CMDSQL + " AND (left(F_CEK,2) IN ('NK','MV','WN','SP','NA','RP','BP','OP','ST')"
CMDSQL = CMDSQL + " or left(F_CEK,3) in ('NBP','PTP','POP','PRE'))"
CMDSQL = CMDSQL + " And agent in (select userid from usertbl where "
CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')"
CMDSQL = CMDSQL + " AND custid in (Select distinct custid from mgm_hst"
CMDSQL = CMDSQL + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT,f_cek "

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
DoEvents
WaitSecs (2)
Set m_hst = New ADODB.Recordset
m_hst.CursorLocation = adUseClient
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        Select Case UCase(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 2))
            Case "NK"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "MV"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "WN"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "BP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "PT"
                If m_hst!F_CEK = "PTP-PA" Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 6)
                ElseIf m_hst!F_CEK = "PTP-PO" Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 6)
                ElseIf m_hst!F_CEK = "PTP-NE" Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 6)
                ElseIf m_hst!F_CEK = "PTP-PR" Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 6)
                ElseIf m_hst!F_CEK = "PTP" Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
                End If
            Case "RP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            Case "NA"
                If Len(m_hst!F_CEK) = 5 Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
                ElseIf Len(m_hst!F_CEK) = 4 Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
                End If
            Case "SP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
              Case "PO"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
             Case "OP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
              '  JMLANTO = JMLANTO
            Case "NB"
                STATUS = Trim(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3))
            Case "PR"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case Else
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            
'            Case "NK"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "MV"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "WN"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "BP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "PT"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "RP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'            Case "NA"
'                If Len(m_hst!F_CEK) = 5 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3) & "M"
'                ElseIf Len(m_hst!F_CEK) = 4 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'                End If
'            Case "SP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'              Case "PO"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'             Case "OP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "NB"
'                STATUS = Trim(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3))
'            Case Else
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            
'             Case "NK"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "MV"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "WN"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "BP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "PT"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "RP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'            Case "NA"
'                'If Len(m_hst!F_CEK) = 5 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'                'ElseIf Len(m_hst!F_CEK) = 4 Then
'                    'STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'                'End If
'            Case "SP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "PO"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "OP"
'                If Len(m_hst!F_CEK) = 2 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 2) + "-"
'                Else
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'                End If
'            Case "NB"
'                STATUS = Trim(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3))
'            Case Else
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            
        End Select
        CMDSQL = CMDSQL + "[" + STATUS + "]"
        CMDSQL = CMDSQL + "=[" + STATUS + "]+" + CStr((IIf(IsNull(m_hst!Jml), "0", m_hst!Jml)))
        If Left(STATUS, 3) = "PTP" Then
        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!F_CEK) Then
        Else
            If m_hst!F_CEK = "" Then
            Else
            
                If m_hst!Jml = 0 Then
                Else
                    M_RPTCONN.Execute CMDSQL
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
LAgent = Empty
CMDSQL = Empty
Exit Sub
'eddder:
 '  If Err.Number = -2147217871 Then
  '      m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
   '     If m_msgbox = vbRetry Then
    '        WaitSecs (5)
      '      Resume
     '   End If
    'Else
    '    MsgBox Err.Description & "eror ambil data yang di pu"
    'End If
    'Resume Next
End Sub

Private Sub AmbilDtYgDiFU_PerAgent_collec()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim m_msgbox As Variant
On Error GoTo eddder

tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
If Option1(0).Value Then

CMDSQL = "select agent,sum(jml) as total FROM (SELECT AGENT, F_CEK, COUNT(F_CEK) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,F_CEK, agent,ttlptp from mgm"
CMDSQL = CMDSQL + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
'CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
CMDSQL = CMDSQL + " AND (left(F_CEK,2) IN ('NK','MV','WN','SP','NA','RP','BP','OP','ST')"
CMDSQL = CMDSQL + " or left(F_CEK,3) in ('NBP','PTP','POP'))"
CMDSQL = CMDSQL + " AND custid in (Select distinct custid from mgm_hst"
CMDSQL = CMDSQL + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT,f_cek) a GROUP BY agent"
Else
If Option1(1).Value Then

CMDSQL = "select agent,sum(jml) as total FROM (SELECT AGENT, F_CEK, COUNT(F_CEK) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,F_CEK, agent,ttlptp from mgm"
CMDSQL = CMDSQL + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
'CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
'CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
CMDSQL = CMDSQL + " AND (left(F_CEK,2) IN ('NK','MV','WN','SP','NA','RP','BP','OP','ST')"
CMDSQL = CMDSQL + " or left(F_CEK,3) in ('NBP','PTP','POP'))"
CMDSQL = CMDSQL + " And agent in (select userid from usertbl where "
CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')"
CMDSQL = CMDSQL + " AND custid in (Select distinct custid from mgm_hst"
CMDSQL = CMDSQL + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT,f_cek) a GROUP BY agent"

End If
End If
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        Jml = IIf(IsNull(m_hst!total), 0, m_hst!total)
        CMDSQL = "Update TrackingRptPerPrgBatch Set VolPayment_LM='" & Jml & "'"
        CMDSQL = CMDSQL + " where AOC='" & LAgent & "'"
        M_RPTCONN.Execute CMDSQL
    m_hst.MoveNext
Wend
Set m_hst = Nothing
LAgent = Empty
CMDSQL = Empty
Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next
End Sub


Private Sub AmbilDtYgDiFU_PerAgent_TARIK()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim m_msgbox As Variant
On Error GoTo eddder

tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
If Option1(0).Value Then

CMDSQL = "SELECT AGENT, F_CEK, COUNT(F_CEK) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,F_CEK, agent,ttlptp from mgmTARIK"
CMDSQL = CMDSQL + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
CMDSQL = CMDSQL + " AND (left(F_CEK,2) IN ('NK','MV','WN','SP','NA','RP','BP','OP','ST')"
CMDSQL = CMDSQL + " or left(F_CEK,3) in ('NBP','PTP','POP'))"
CMDSQL = CMDSQL + " AND custid in (Select distinct custid from mgm_hst"
CMDSQL = CMDSQL + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT,f_cek"

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

CMDSQL = "SELECT AGENT, F_CEK, COUNT(F_CEK) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,F_CEK, agent,ttlptp from mgmTARIK"
CMDSQL = CMDSQL + " where tglcall between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " And nextactdate between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
'CMDSQL = CMDSQL + " and agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "'"
CMDSQL = CMDSQL + " AND (left(F_CEK,2) IN ('NK','MV','WN','SP','NA','RP','BP','OP','ST')"
CMDSQL = CMDSQL + " or left(F_CEK,3) in ('NBP','PTP','POP'))"
CMDSQL = CMDSQL + " And agent in (select userid from usertbl where "
CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')"
CMDSQL = CMDSQL + " AND custid in (Select distinct custid from mgm_hst"
CMDSQL = CMDSQL + " where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT,f_cek"

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
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        Select Case UCase(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 2))
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
                If Len(m_hst!F_CEK) = 5 Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
                ElseIf Len(m_hst!F_CEK) = 4 Then
                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
                End If
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
            
'            Case "NK"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "MV"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "WN"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "BP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "PT"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "RP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'            Case "NA"
'                If Len(m_hst!F_CEK) = 5 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3) & "M"
'                ElseIf Len(m_hst!F_CEK) = 4 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'                End If
'            Case "SP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'              Case "PO"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'             Case "OP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "NB"
'                STATUS = Trim(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3))
'            Case Else
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            
'             Case "NK"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "MV"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "WN"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
'            Case "BP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "PT"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "RP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'            Case "NA"
'                'If Len(m_hst!F_CEK) = 5 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'                'ElseIf Len(m_hst!F_CEK) = 4 Then
'                    'STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
'                'End If
'            Case "SP"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "PO"
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'            Case "OP"
'                If Len(m_hst!F_CEK) = 2 Then
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 2) + "-"
'                Else
'                    STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
'                End If
'            Case "NB"
'                STATUS = Trim(Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3))
'            Case Else
'                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            
        End Select
        CMDSQL = CMDSQL + "[" + STATUS + "]"
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_hst!Jml), 0, m_hst!Jml)) + " "
        If STATUS = "PTP" Then
        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!F_CEK) Then
        Else
            If m_hst!F_CEK = "" Then
            Else
            
                If m_hst!Jml = 0 Then
                Else
                    M_RPTCONN.Execute CMDSQL
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
LAgent = Empty
CMDSQL = Empty
Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next
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
CMDSQL = " SELECT AGENT, StatusCall, COUNT(StatusCall) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,StatusCall, agent,ttlptp from mgm"
CMDSQL = CMDSQL + " where tglcall >= '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and tglcall <= '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource >= '" + Combo1(0).Text + "' and recsource <= '" + Combo1(1).Text + "'"
'CMDSQL = CMDSQL + " And Kethslkerja <> 'I') "
CMDSQL = CMDSQL + " And custid in (Select distinct custid from mgm_hst"
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
CMDSQL = CMDSQL + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT, STATUSCALL"

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
CMDSQL = " SELECT AGENT, StatusCall, COUNT(StatusCall) AS Jml,sum(ttlptp) as jmlPTP FROM"
CMDSQL = CMDSQL + " (select custid, recsource,StatusCall, agent,ttlptp from mgm"
CMDSQL = CMDSQL + " where tglcall >= '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and tglcall <= '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and recsource >= '" + Combo1(0).Text + "' and recsource <= '" + Combo1(1).Text + "'"
'CMDSQL = CMDSQL + " And Kethslkerja <> 'I') "
CMDSQL = CMDSQL + " And agent in (select userid from usertbl where "
CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')"
CMDSQL = CMDSQL + " And custid in (Select distinct custid from mgm_hst"
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
CMDSQL = CMDSQL + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " A GROUP BY AGENT, STATUSCALL"


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
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
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
        CMDSQL = CMDSQL + "[" + STATUS + "]"
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_hst!Jml), 0, m_hst!Jml)) + " "
        If STATUS = "PTP" Then
        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!statuscall) Then
        Else
            If m_hst!statuscall = "" Then
            Else
            
                If m_hst!Jml = 0 Then
                Else
                    M_RPTCONN.Execute CMDSQL
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
LAgent = Empty
CMDSQL = Empty
Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
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
CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ) "
CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "AND mgm_hst.F_CEK IN ('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
Else
If Option1(1).Value Then
CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent in (select userid from usertbl where "
CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "'))"
CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "AND mgm_hst.F_CEK IN ('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent "
End If
End If
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
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
        STATUS = Left(STATUS, 3) + Last
        CMDSQL = CMDSQL + "[" + STATUS + "]"
        CMDSQL = CMDSQL + "=[" + STATUS + "]+ " + CStr(IIf(IsNull(m_hst!Jml), 0, m_hst!Jml)) + " "
        If STATUS = "PTP_LM" Then
        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!F_CEK) Then
        Else
            If m_hst!F_CEK = "" Then
            Else
            
                If m_hst!Jml = 0 Then
                Else
                    M_RPTCONN.Execute CMDSQL
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
LAgent = Empty
CMDSQL = Empty
Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
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
CMDSQL = "select  tblvisit.ffc, tblvisit.StatusVisit,count(tblvisit.statusVisit) as jml from tblvisit "
CMDSQL = CMDSQL + " inner join (SELECT custid, ffc, max(RequestDate) as tglmax from tblvisit where "
CMDSQL = CMDSQL + " RequestDate Between  '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
CMDSQL = CMDSQL + " and ffc in (select userid from usertbl where userid >= '" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' )group by custid,ffc)  as  a "
CMDSQL = CMDSQL + " on tblvisit.custid = a.custid and tblvisit.requestdate=a.tglmax "
CMDSQL = CMDSQL + " inner join mgm on mgm.custid= a.custid where tblvisit.statusvisit in('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'  "
CMDSQL = CMDSQL + "group by tblvisit.ffc, statusVisit "
Else
If Option1(1).Value Then
CMDSQL = "select  tblvisit.ffc, tblvisit.StatusVisit,count(tblvisit.statusVisit) as jml from tblvisit "
CMDSQL = CMDSQL + " inner join (SELECT custid, ffc, max(RequestDate) as tglmax from tblvisit where "
CMDSQL = CMDSQL + " RequestDate Between  '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
CMDSQL = CMDSQL + " and ffc in (select userid from usertbl where SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' )group by custid,ffc)  as  a "
CMDSQL = CMDSQL + " on tblvisit.custid = a.custid and tblvisit.requestdate=a.tglmax "
CMDSQL = CMDSQL + " inner join mgm on mgm.custid= a.custid where tblvisit.statusvisit in('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC','MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G','NA-H','NA-P','NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J','RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP') "
CMDSQL = CMDSQL + " and recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'  "
CMDSQL = CMDSQL + "group by tblvisit.ffc, statusVisit "
End If
End If
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = Trim(CStr(IIf(IsNull(m_hst!FFC), "", m_hst!FFC)))
        CMDSQL = "Update TrackingRptField Set "
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
        CMDSQL = CMDSQL + "[" + STATUS + "]"
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_hst!Jml), 0, m_hst!Jml)) + " "
'        If STATUS = "PTP" Then
'        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
'        End If
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!StatusVisit) Then
        Else
            If m_hst!StatusVisit = "" Then
            Else
            
                If m_hst!Jml = 0 Then
                Else
                    M_RPTCONN.Execute CMDSQL
               End If
        End If
       End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
STATUS = Empty
CMDSQL = Empty
LAgent = Empty
Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox Err.Description
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
CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where MONTH(tgl) ='" + VarMonth + "' AND YEAR(tgl)='" + VarYear + "' AND "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ) "
CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
Else
If Option1(1).Value Then
CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where MONTH(tgl) ='" + VarMonth + "' AND YEAR(tgl)='" + VarYear + "' AND "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent in (select userid from usertbl where "
CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "'))"
CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
 End If
End If
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
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
        CMDSQL = CMDSQL + "[" + STATUS + LastMonth + "]"
'        CMDSQL = CMDSQL + "[" + STATUS + "]"
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_hst!Jml), 0, m_hst!Jml)) + " "
        If STATUS = "PTP" Then
        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!F_CEK) Then
        Else
            If m_hst!F_CEK = "" Then
            Else
                If m_hst!Jml = 0 Then
                Else
                   
                    M_RPTCONN.Execute CMDSQL
                End If
            End If
        End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing

Exit Sub
eddder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
  '      MsgBox Err.Description
    End If
    Resume Next

End Sub

Private Sub Check1_Click(Index As Integer)
Dim CMDSQL2 As String
Combo1(0).Clear
Combo1(1).Clear
Combo2(0).Clear
Combo2(1).Clear
Check2(0).Value = 0
Check2(1).Value = 0
Select Case Index
Case 0
        CMDSQL = "SELECT * FROM usertbl WHERE AKTIF = 0 AND USERID IN ("
        CMDSQL = CMDSQL + "SELECT AGENT FROM mgm WHERE RECSOURCE LIKE '%CF%' AND RECSOURCE NOT LIKE '%CFR%') ORDER BY USERID"

        Check1(1).Value = 0
        Check1(2).Value = 0
Case 1
        CMDSQL = "SELECT * FROM usertbl WHERE AKTIF = 0 AND USERID IN ("
        CMDSQL = CMDSQL + "SELECT AGENT FROM mgm WHERE RECSOURCE LIKE '%CFR%') ORDER BY USERID"
        
        Check1(0).Value = 0
        Check1(2).Value = 0
Case 2
        CMDSQL = "SELECT * FROM usertbl WHERE AKTIF = 0 AND USERID IN ("
        CMDSQL = CMDSQL + "SELECT AGENT FROM mgm WHERE RECSOURCE LIKE '%PIL%' AND RECSOURCE NOT LIKE '%CFR%') ORDER BY USERID"
        
        Check1(1).Value = 0
        Check1(0).Value = 0
End Select


Dim M_OBJRS As ADODB.Recordset
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJRS.EOF
      Combo2(0).AddItem M_OBJRS!USERID
      Combo2(1).AddItem M_OBJRS!USERID
M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing

End Sub

Private Sub Check2_Click(Index As Integer)
Dim CMDSQL2 As String
Dim M_OBJRS As ADODB.Recordset
Combo1(0).Clear
Combo1(1).Clear

Select Case Index
Case 0
    If Check1(0).Value = 1 Then
        CMDSQL2 = "SELECT * FROM DATASOURCETBL WHERE KODEDS LIKE '%CF%' AND KODEDS NOT LIKE '%CFR%' ORDER BY KODEDS"
    ElseIf Check1(1).Value = 1 Then
        CMDSQL2 = "SELECT * FROM DATASOURCETBL WHERE KODEDS LIKE '%CFR%' ORDER BY KODEDS"
    ElseIf Check1(2).Value = 1 Then
        CMDSQL2 = "SELECT * FROM DATASOURCETBL WHERE KODEDS LIKE '%PIL%' ORDER BY KODEDS"
    Else
        CMDSQL2 = "SELECT * FROM DATASOURCETBL ORDER BY KODEDS"
    End If
    Check2(1).Value = 0
    dtcek = False
Case 1
    If Check1(0).Value = 1 Then
        CMDSQL2 = "SELECT * FROM DATASOURCETBLTARIK WHERE KODEDS LIKE '%CF%' AND KODEDS NOT LIKE '%CFR%' ORDER BY KODEDS"
    ElseIf Check1(1).Value = 1 Then
        CMDSQL2 = "SELECT * FROM DATASOURCETBLTARIK WHERE KODEDS LIKE '%CFR%' ORDER BY KODEDS"
    ElseIf Check1(2).Value = 1 Then
        CMDSQL2 = "SELECT * FROM DATASOURCETBLTARIK WHERE KODEDS LIKE '%PIL%' ORDER BY KODEDS"
    Else
        CMDSQL2 = "SELECT * FROM DATASOURCETBLTARIK ORDER BY KODEDS"
    End If
    Check2(0).Value = 0
    dtcek = True
End Select

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJRS.EOF
    Combo1(0).AddItem M_OBJRS!KODEDS
    Combo1(1).AddItem M_OBJRS!KODEDS
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
    Call Combo1_LostFocus(Index)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_OBJRS As New ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Combo1_LostFocusErr
    M_OBJRS.CursorLocation = adUseClient
    If dtcek = True Then
        CMDSQL = "Select * from datasourcetbltarik where kodeds ='" + Combo1(Index).Text + "'"
    Else
        CMDSQL = "Select * from datasourcetbl where kodeds ='" + Combo1(Index).Text + "'"
    End If
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not M_OBJRS.EOF Then
        Combo1(Index).Text = M_OBJRS!KODEDS
    Else
        Combo1(Index).Text = Empty
    End If
Exit Sub
Combo1_LostFocusErr:
    MsgBox Err.Description
End Sub

Private Sub Combo2_Click(Index As Integer)
    Call Combo2_LostFocus(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim M_OBJRS As New ADODB.Recordset
On Error GoTo Combo2_LostFocusErr
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND USERID ='" + Combo2(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not M_OBJRS.EOF Then
        Combo2(Index).Text = M_OBJRS!USERID
    Else
        Combo2(Index).Text = Empty
    End If
Exit Sub
Combo2_LostFocusErr:
    MsgBox Err.Description
End Sub
Private Sub hitung_JmlData_PerAgent_PTP()
Dim M_OBJRS As New ADODB.Recordset
Dim PTP As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient

'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "Select Agent, sum(ttlptp) as JMLVOL, count(f_cek) as PTP from mgm  where F_CEK ='PTP' AND recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND Tglincoming BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    'JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "0", m_objrs!jml))
    JUMLAHVOL = CStr(IIf(IsNull(M_OBJRS!JMLVOL), "0", M_OBJRS!JMLVOL))
    LAgent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
    PTP = CStr(IIf(IsNull(M_OBJRS!PTP), "0", M_OBJRS!PTP))
    CMDSQL = "Update TrackingRptPerPrgBatch set  VOLPTP1 = " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    CMDSQL = "Update TrackingRptPerPrgBatch set  PTP1 = " + PTP + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAHVOL = Empty
LAgent = Empty
PTP = Empty
CMDSQL = Empty


Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub hitung_JmlData_PerAgent_PTP_TARIK()
Dim M_OBJRS As New ADODB.Recordset
Dim PTP As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient

'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "Select Agent, sum(ttlptp) as JMLVOL, count(f_cek) as PTP from mgmTARIK where F_CEK ='PTP' AND recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND Tglincoming BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    'JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "0", m_objrs!jml))
    JUMLAHVOL = CStr(IIf(IsNull(M_OBJRS!JMLVOL), "0", M_OBJRS!JMLVOL))
    LAgent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
    PTP = CStr(IIf(IsNull(M_OBJRS!PTP), "0", M_OBJRS!PTP))
    CMDSQL = "Update TrackingRptPerPrgBatch set  VOLPTP1 = " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    CMDSQL = "Update TrackingRptPerPrgBatch set  PTP1 = " + PTP + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAHVOL = Empty
LAgent = Empty
PTP = Empty
CMDSQL = Empty


Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
'On Error GoTo Command1_ClickeR
If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
    TDBDate1(0).Value = "01/01/1990"
    TDBDate1(1).Value = "31/12/2020"
End If

If Combo1(0).Text = Empty And Combo1(1).Text = Empty Then
    Combo1(0).Text = "-----"
    Combo1(1).Text = "ZZZZZ"
End If
If Option1(0).Value = False And Option1(1).Value = False Then
If Combo2(0).Text = Empty And Combo2(1).Text = Empty Then
    Combo2(0).Text = "-----"
    Combo2(1).Text = "ZZZZZ"
End If
End If
ProgressBar1.Visible = True

Select Case Index
    Case 0
    Select Case listview1.SelectedItem.Text
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
'            'RPT.ReportFileName = App.Path + "\Report\Tracking ReportSPV_sum.rpt"
'            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\Tracking ReportSPV_sum.rpt"
'            Call SHOW_PRN
            
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
'            'RPT.ReportFileName = App.Path + "\Report\Tracking ReportAgent.rpt"
'            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\Tracking ReportAgent.rpt"
'            Call SHOW_PRN
        
       Case 1
            Call Isi_Agent_mgm
            Call hitung_JmlData_PerAgent_mgm
            Call AmbilDtYgDiFU_PerAgent
            Call VolumeUtilized
            Call ReportPTPNego
            'Call AmbilDataYgDiFU_LastMonth
            Call Isi_Settled_Payment
'           Call Isi_Progess_OF_PAyment
            Call hitung_JmlData_PerAgent_PTP
          ' Call Hitung_JmlLeadsPerAgent
           Call Hitung_Vol_PTP
           Call hitung_BatchCallInitilized_PerAgent_mgm
           Call Hitung_Number_of_Payment
           Call Hitung_Volume_of_Payment
           Call Hitung_Payment
           
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\Tracking ReportSPVGlobal.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\Tracking ReportSPVGlobalPIL.rpt"
            Call SHOW_PRN
        
        Case 2
            Call Isi_Agent_mgm
            Call hitung_JmlData_PerAgent_mgm
            Call AmbilDtYgDiFU_PerAgent
            Call VolumeUtilized
          '  Call ReportPTPNego
            'Call AmbilDataYgDiFU_LastMonth
'            Call Isi_Settled_Payment
'            Call Isi_Progess_OF_PAyment
'            Call hitung_JmlData_PerAgent_PTP
          '' Call Hitung_JmlLeadsPerAgent
          '  Call Hitung_Vol_PTP
            Call hitung_BatchCallInitilized_PerAgent_mgm
            Call Hitung_Number_of_Payment
            Call Hitung_Volume_of_Payment
        
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\Tracking ReportagentGlobal.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\Tracking ReportAgentGlobalPIL.rpt"
            Call SHOW_PRN
            
         Case 20
            Call Isi_Agent_mgm_TARIK
            Call hitung_JmlData_PerAgent_mgm_TARIK
            Call AmbilDtYgDiFU_PerAgent_TARIK
            Call VolumeUtilized_TARIK
            Call ReportPTPNego_TARIK
            'Call AmbilDataYgDiFU_LastMonth
            Call Isi_Settled_Payment_TARIK
'            Call Isi_Progess_OF_PAyment
            Call hitung_JmlData_PerAgent_PTP_TARIK
          '' Call Hitung_JmlLeadsPerAgent
            Call Hitung_Vol_PTP_TARIK
            Call hitung_BatchCallInitilized_PerAgent_mgm_TARIK
            Call Hitung_Number_of_Payment_TARIK
            Call Hitung_Volume_of_Payment_TARIK
        
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\Tracking ReportagentGlobal.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\Tracking ReportAgentGlobalPIL.rpt"
            Call SHOW_PRN
        
        Case 3
            Call Isi_Agent_mgm
            Call hitung_JmlData_PerAgent_mgm
            Call AmbilDtYgDiFU_PerAgentcall
            Call VolumeUtilized
            Call hitung_BatchCallInitilized_PerAgent_mgm
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\ReportAgentGlobalstatuscall.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\ReportAgentGlobalstatuscall.rpt"
            Call SHOW_PRN
            
        Case 4
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(3) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            'RPT.ReportFileName = App.Path + "\Report\RptDistribusi.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\RptDistribusi.rpt"
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
'            'RPT.ReportFileName = App.Path + "\Report\historyCall.rpt"
'            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\historyCall.rpt"
'            Call SHOW_PRN
'        Case 7
'            Call TrackingReservedPTP
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
'            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
'            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
'            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
'            'RPT.ReportFileName = App.Path + "\Report\historyCall_custid.rpt"
'            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\TrackingReservedPTP.rpt"
'            Call SHOW_PRN
            
                 
        Case 5
        Call Tracking_PTP_Report
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\RptPromiseToPay.rpt"
            Call SHOW_PRN
         
'         Case 9
'         Call Isi_Report_PTP_lunas
'            WaitSecs (2)
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
'            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
'            'RPT.ReportFileName = App.Path + "\Report\ActualPay.rpt"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ActualPay.rpt"
'            Call SHOW_PRN
            
        Case 6
           Call Isi_Report_PTP_Jatuh_Tempo
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            'RPT.ReportFileName = App.Path + "\Report\RptPTPJatuhTempo.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\RptPTPJatuhTempo.rpt"
            Call SHOW_PRN
            
        Case 7
          Call TrackingReservedPTP
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            'RPT.ReportFileName = App.Path + "\Report\RptPTPJatuhTempo.rpt"
            RPT.ReportFileName = "D:\COLLECTION_AWARNESS\Report\ReservedPTP.rpt"
            Call SHOW_PRN
            
        
        
        Case 21
           Call Isi_Report_PTP_Jatuh_TempoTARIK
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            'RPT.ReportFileName = App.Path + "\Report\RptPTPJatuhTempo.rpt"
            RPT.ReportFileName = "D:\COLLECTION\Report\RptPTPJatuhTempo.rpt"
            Call SHOW_PRN
         Case 22
            Call Isi_Agent_mgm
            Call hitung_JmlData_PerAgent_mgm
            Call Hitung_Number_of_PaymentCOLL
            'Call amount_paid
            Call AMOUNT_COLLECTED
            Call AmbilDtYgDiFU_PerAgent
            'Call AmbilDtYgDiFU_PerAgent_collec
            Call hitung_BatchCallInitilized_PerAgent_mgm
            'Call hitung_jmlPTP
           ' Call PTP_withpayment
           ' Call hitung_jmlPOP_LM
            Call hitung_jmlPOP
            Call HITUNGHACCOUNTPTP
            Call hitungcountpop
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\Tracking ReportagentGlobal.rpt"
            RPT.ReportFileName = "D:\COLLECTION\Report\Tracking Reportcollectornew16baru.rpt"
            Call SHOW_PRN
            
'         Case 11
'         Call Isi_Report_FormVisit
'            WaitSecs (2)
'            RPT.Reset
''            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
''            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
''            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
''            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
''            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
'            'RPT.ReportFileName = App.Path + "\Report\RptVisit.rpt"
'            RPT.ReportFileName = "D:\COLLECTION\Report\RptVisit.rpt"
'            Call SHOW_PRN
'
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
'            'RPT.ReportFileName = App.Path + "\Report\RptTrackingField.rpt"
'            RPT.ReportFileName = "D:\COLLECTION\Report\RptTrackingField.rpt"
'            Call SHOW_PRN
     
     'Case 13
'            'WaitSecs (2)
'            'Call Isi_Agent_mgm
'            'Call hitung_BatchCallInitilized_PerAgent_mgm
'            Call hitung_BatchCallInitilized_PerAgent_Compare
'            Call AmbilDtYgDiFU_PerAgent
'            Call AmbilDtYgDiFU_PerAgent_Compare
'            RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
'            'RPT.ReportFileName = App.Path + "\Report\ChartUtilizedCallAgent.rpt"
'            RPT.ReportFileName = "D:\COLLECTION\Report\ChartUtilizedCallAgent.rpt"
'            Call SHOW_PRN
'
     Case 14
            WaitSecs (2)
            Call Isi_Agent_mgm
            Call AmbilDtYgDiFU_PerAgent
            Call AmbilDtYgDiFU_PerAgent_Compare
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\ChartUtilized.rpt"
            RPT.ReportFileName = "D:\COLLECTION\Report\ChartUtilized.rpt"
            Call SHOW_PRN
            
     Case 15
            WaitSecs (2)
            Call Isi_Agent_mgm
            Call Hitung_Volume_of_Payment
            Call Hitung_Volume_of_Payment_Compare
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\ChartPayment.rpt"
            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPayment.rpt"
            Call SHOW_PRN
     
     Case 16
            WaitSecs (2)
            Call Isi_Agent_mgm
            Call ReportPTPNego
            Call ReportPTPNego_Compare
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\ChartPTP.rpt"
            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPTP.rpt"
            Call SHOW_PRN
    Case 17
            WaitSecs (2)
            Call Isi_Agent_mgm
            Call ReportPTPNego
            Call ReportPTPNego_Compare
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\ChartPTP.rpt"
            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPTP.rpt"
            Call SHOW_PRN
        Case 18
            WaitSecs (2)
            Call Isi_Agent_mgm
            Call ReportPTPNego
            Call ReportPTPNego_Compare
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglShow2 = totext('" + CStr(TDBDate1(2).Text & " " & DTimeLastCall(2).Text) + "')"
            RPT.Formulas(5) = "@TglShow3 = totext('" + CStr(TDBDate1(3).Text & " " & DTimeLastCall(3).Text) + "')"
            'RPT.ReportFileName = App.Path + "\Report\ChartPTP.rpt"
            RPT.ReportFileName = "D:\COLLECTION\Report\ChartPTP.rpt"
            Call SHOW_PRN
            
            
        Case 8
            '@@ Report POP BP1 Ritpill [C.24-11-09] -- POSTGREE
            Call ISI_DATA_POP_BP1
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPOPBP.rpt"
            Call SHOW_PRN
            
       '@@ Report POP BP2 Ritpill [C.24-11-09] -- POSTGREE
        Case 9
            Call ISI_DATA_POP_BP2
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPOPBP2.rpt"
            Call SHOW_PRN
        
        '@@ Report POP BP3 Ritpill [C.24-11-09] -- POSTGREE
        Case 10
            Call ISI_DATA_POP_BP3
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPOPBP3.rpt"
            Call SHOW_PRN
             
        '@@ Report BP1 Ritpill [C.24-11-09] -- POSTGREE
        Case 11
            Call ISI_DATA_BP1
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptBP1.rpt"
            Call SHOW_PRN
        
        '@@ Report BP2 Ritpill [C.24-11-09] -- POSTGREE
        Case 12
            Call ISI_DATA_BP2
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptBP2.rpt"
            Call SHOW_PRN
        
        '@@ Report BP3 Ritpill [C.24-11-09] -- POSTGREE
        Case 13
            Call ISI_DATA_BP3
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            RPT.Formulas(6) = "@spvawal = totext('" + CStr(Combo3(0).Text) + "')"
            RPT.Formulas(7) = "@spvakhir = totext('" + CStr(Combo3(1).Text) + "')"
            'RPT.ReportFileName = "D:\COLLECTION\Report\RptPromiseToPay.rpt"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptBP3.rpt"
            Call SHOW_PRN
            
    End Select
    ProgressBar1.Visible = False
    Case 1
        Unload Me
End Select
ProgressBar1.Visible = False



Exit Sub
'Command1_ClickeR:
 '   MsgBox Err.Description
'    Resume
End Sub

Private Sub amount_paid()
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "select agent,count(custid) as jml from vwlunas where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND paydate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set VOLPAYMENT =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next

End Sub

Private Sub AMOUNT_COLLECTED()
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "select agent,sum(Payment) as jml from tbllunas where paydate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set NPAYMENT =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next

End Sub

Private Sub Isi_Settled_Payment()
Dim M_OBJRS As New ADODB.Recordset
Dim LRECSOURCE As String
Dim m_msgbox As Variant


On Error GoTo hitung_JmlDataer
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, count(tgllunas) as jml from mgm where tgllunas BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set SETTLED_PAYMENT =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    CMDSQL = Empty
    LAgent = Empty
    JUMLAH = Empty
    
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub

Private Sub Isi_Settled_Payment_TARIK()
Dim M_OBJRS As New ADODB.Recordset
Dim LRECSOURCE As String
Dim m_msgbox As Variant


On Error GoTo hitung_JmlDataer
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, count(tgllunas) as jml from mgmTARIK where tgllunas BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set SETTLED_PAYMENT =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    CMDSQL = Empty
    LAgent = Empty
    JUMLAH = Empty
    
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub


Private Sub Isi_Progess_OF_PAyment()
Dim M_OBJRS As New ADODB.Recordset
Dim LRECSOURCE As String
Dim m_msgbox As Variant


On Error GoTo hitung_JmlDataer
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, count(tglPOP) as jml from mgm where tglPOP BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch set POP2 =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    JUMLAH = Empty
    LAgent = Empty
    CMDSQL = Empty
    
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub


Private Sub Hitung_Number_of_Payment()
Dim M_OBJRS As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, count(custid) as jml from (select distinct custid,agent from HtgNumberOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"

   ' CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where F_CEK ='PTP' AND recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND TGLINCOMING BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set NPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub

Private Sub Hitung_Number_of_PaymentCOLL()
Dim M_OBJRS As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, count(custid) as jml from (select distinct custid,agent from HtgNumberOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"

   ' CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where F_CEK ='PTP' AND recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND TGLINCOMING BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set VOLPAYMENT =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub


Private Sub Hitung_Number_of_Payment_TARIK()
Dim M_OBJRS As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, count(custid) as jml from (select distinct custid,agent from HtgNumberOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"

   ' CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where F_CEK ='PTP' AND recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND TGLINCOMING BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set NPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub

Private Sub Hitung_Volume_of_Payment()
Dim M_OBJRS As New ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, sum(Payment) as jml from (select * from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set VolPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    JUMLAH = Empty
    LAgent = Empty
    CMDSQL = Empty
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If
End Sub

Private Sub Hitung_Volume_of_Payment_TARIK()
Dim M_OBJRS As New ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, sum(Payment) as jml from (select * from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set VolPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    JUMLAH = Empty
    LAgent = Empty
    CMDSQL = Empty
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If
End Sub

Private Sub Hitung_Volume_of_Payment_Compare()
Dim M_OBJRS As New ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient

CMDSQL = "select agent, sum(Payment) as jml from (select * from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' AND '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set VolPayment_LM =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    JUMLAH = Empty
    LAgent = Empty
    CMDSQL = Empty
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If
End Sub

Private Sub Hitung_Vol_PTP()
Dim M_OBJRS As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim LRECSOURCE As String
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient


If Option1(0).Value Then
CMDSQL = "SELECT mgm_hst.agent, mgm_hst.f_cek,count(mgm_hst.f_cek) as JML,  sum(mgm.ttlptp) as jmlPTP from mgm_hst INNER JOIN (SELECT custid, max(tgl) as tglmax FROM mgm_hst "
CMDSQL = CMDSQL + "where tgl BETWEEN ' 1990-01-01 12:00:00 AM' and '2020-12-31 11:59:00 PM'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent between  '" + Combo2(0).Text + "' and  '" + Combo2(1).Text + "') and "
CMDSQL = CMDSQL + " f_cek='PTP' group by custid) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax inner join mgm on mgm.custid = a.custid and "
CMDSQL = CMDSQL + "mgm_hst.agent=mgm.agent where recsource BETWEEN'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent, mgm_hst.f_cek "
CMDSQL = CMDSQL + "order by mgm_hst.agent"
Else
If Option1(1).Value Then
CMDSQL = "SELECT mgm_hst.agent, mgm_hst.f_cek,count(mgm_hst.f_cek) as JML, sum(mgm.ttlptp) as jmlPTP from mgm_hst INNER JOIN (SELECT custid, max(tgl) as tglmax FROM mgm_hst "
CMDSQL = CMDSQL + "where tgl BETWEEN ' 1990-01-01 12:00:00 AM' and '2020-12-31 11:59:00 PM'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent in (select userid from usertbl where SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')) and "
CMDSQL = CMDSQL + " f_cek='PTP' group by custid) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax inner join mgm on mgm.custid = a.custid and "
CMDSQL = CMDSQL + "mgm_hst.agent=mgm.agent where recsource BETWEEN'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent, mgm_hst.f_cek "
CMDSQL = CMDSQL + "order by mgm_hst.agent"
End If
End If
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!jmlPTP), "0", M_OBJRS!jmlPTP))
        LAgent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch set VolPTP =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub

Private Sub Hitung_Vol_PTP_TARIK()
Dim M_OBJRS As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim LRECSOURCE As String
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient


If Option1(0).Value Then
CMDSQL = "SELECT mgm_hst.agent, mgm_hst.f_cek,count(mgm_hst.f_cek) as JML,  sum(mgmTARIK.ttlptp) as jmlPTP from mgm_hst INNER JOIN (SELECT custid, max(tgl) as tglmax FROM mgm_hst "
CMDSQL = CMDSQL + "where tgl BETWEEN ' 1990-01-01 12:00:00 AM' and '2020-12-31 11:59:00 PM'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgmTARIK where agent between  '" + Combo2(0).Text + "' and  '" + Combo2(1).Text + "') and "
CMDSQL = CMDSQL + " f_cek='PTP' group by custid) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax inner join mgmTARIK on mgmTARIK.custid = a.custid and "
CMDSQL = CMDSQL + "mgm_hst.agent=mgmTARIK.agent where recsource BETWEEN'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent, mgm_hst.f_cek "
CMDSQL = CMDSQL + "order by mgm_hst.agent"
Else
If Option1(1).Value Then
CMDSQL = "SELECT mgm_hst.agent, mgm_hst.f_cek,count(mgm_hst.f_cek) as JML, sum(mgmTARIK.ttlptp) as jmlPTP from mgm_hst INNER JOIN (SELECT custid, max(tgl) as tglmax FROM mgm_hst "
CMDSQL = CMDSQL + "where tgl BETWEEN ' 1990-01-01 12:00:00 AM' and '2020-12-31 11:59:00 PM'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgmTARIK where agent in (select userid from usertbl where SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')) and "
CMDSQL = CMDSQL + " f_cek='PTP' group by custid) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax inner join mgmTARIK on mgmTARIK.custid = a.custid and "
CMDSQL = CMDSQL + "mgm_hst.agent=mgmTARIK.agent where recsource BETWEEN'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
CMDSQL = CMDSQL + "group by mgm_hst.agent, mgm_hst.f_cek "
CMDSQL = CMDSQL + "order by mgm_hst.agent"
End If
End If
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!jmlPTP), "0", M_OBJRS!jmlPTP))
        LAgent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch set VolPTP =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If

End Sub

Private Sub Isi_Report_PTP_Jatuh_TempoTARIK()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingPTP"
Set M_OBJRS = New ADODB.Recordset
Set M_OBJRS1 = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient
If Option1(1).Value = True Then
CMDSQL = "Select * from reportPTPTARIK where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') ORDER BY AGENT"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
       CMDSQL = "Select * from reportPTPTARIK where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent between '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ORDER BY AGENT"
        M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
M_OBJRS1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_OBJRS.RecordCount + 1
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    M_OBJRS1.AddNew
    M_OBJRS1!agent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
    M_OBJRS1!CustId = CStr(IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId))
    M_OBJRS1!Name = CStr(IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name))
    M_OBJRS1!TglPTP = CStr(IIf(IsNull(M_OBJRS!PromiseDate), "2020-12-30", M_OBJRS!PromiseDate))
    M_OBJRS1!inputdate = CStr(IIf(IsNull(M_OBJRS!inputdate), "2020-12-30", M_OBJRS!inputdate))
    M_OBJRS1!ttlptp = CStr(IIf(IsNull(M_OBJRS!PromisePay), "0", M_OBJRS!PromisePay))
    M_OBJRS1!amountwo = CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo))
    M_OBJRS1!Principle = CStr(IIf(IsNull(M_OBJRS!Principal), "0", M_OBJRS!Principal))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
    'Resume

End Sub

Private Sub Isi_Report_PTP_Jatuh_Tempo()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingPTP"
Set M_OBJRS = New ADODB.Recordset
Set M_OBJRS1 = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient
If Option1(1).Value = True Then
CMDSQL = "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') ORDER BY AGENT"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
       CMDSQL = "Select * from reportPTP where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent between '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ORDER BY AGENT"
        M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
M_OBJRS1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_OBJRS.RecordCount + 1
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    M_OBJRS1.AddNew
    M_OBJRS1!agent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
    M_OBJRS1!CustId = CStr(IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId))
    M_OBJRS1!Name = CStr(IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name))
    M_OBJRS1!TglPTP = CStr(IIf(IsNull(M_OBJRS!PromiseDate), "2020-12-30", M_OBJRS!PromiseDate))
    M_OBJRS1!inputdate = CStr(IIf(IsNull(M_OBJRS!inputdate), "2020-12-30", M_OBJRS!inputdate))
    M_OBJRS1!ttlptp = CStr(IIf(IsNull(M_OBJRS!PromisePay), "0", M_OBJRS!PromisePay))
    M_OBJRS1!amountwo = CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo))
    M_OBJRS1!Principle = CStr(IIf(IsNull(M_OBJRS!Principal), "0", M_OBJRS!Principal))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
    'Resume

End Sub

Private Sub Isi_Report_FormVisit()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from FormVisit"
Set M_OBJRS = New ADODB.Recordset
Set M_OBJRS1 = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient
If Option1(0).Value = True Then
CMDSQL = "SELECT TblVisit.*, mgm.Principal AS PRINCIPLE, mgm.AmountWo AS AmountWO,mgm.name as NAME "
CMDSQL = CMDSQL + "FROM mgm INNER JOIN "
CMDSQL = CMDSQL + "TblVisit ON dbo.mgm.CUSTID = dbo.TblVisit.CUSTID "
CMDSQL = CMDSQL + "WHERE TblVisit.agent between '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' "
CMDSQL = CMDSQL + "AND tblvisit.requestDate between  '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + " AND sts = '0'"
CMDSQL = CMDSQL + " ORDER BY tblvisit.VisitNo"
'CMDSQL = "Select * from mgm where f_cek='PTP' and tdbdatePTP between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') ORDER BY AGENT"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(1).Value = True Then
CMDSQL = "SELECT TblVisit.*, mgm.Principal AS PRINCIPLE, mgm.AmountWo AS AmountWO, mgm.name as NAME "
CMDSQL = CMDSQL + "FROM mgm INNER JOIN "
CMDSQL = CMDSQL + "TblVisit ON dbo.mgm.CUSTID = dbo.TblVisit.CUSTID "
CMDSQL = CMDSQL + "WHERE TblVisit.agent in (SELECT userid from usertbl where SPVCODE  >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') "
CMDSQL = CMDSQL + "AND tblvisit.requestDate between  '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + " AND sts = '0'"
CMDSQL = CMDSQL + " ORDER BY tblvisit.VisitNo"
       M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
M_OBJRS1.Open "Select * from formVisit", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_OBJRS.RecordCount + 1
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    M_OBJRS1.AddNew
    M_OBJRS1!agent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    M_OBJRS1!FFC = Trim(CStr(IIf(IsNull(M_OBJRS!FFC), "", M_OBJRS!FFC)))
    M_OBJRS1!CustId = Trim(CStr(IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId)))
    M_OBJRS1!Name = Trim(CStr(IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name)))
    M_OBJRS1!NoVisit = Trim(CStr(IIf(IsNull(M_OBJRS!VisitNo), "", M_OBJRS!VisitNo)))
    M_OBJRS1!RequestDate = Trim(CStr(IIf(IsNull(M_OBJRS!RequestDate), "2020-12-30", M_OBJRS!RequestDate)))
    M_OBJRS1!DetailsR = Trim(CStr(IIf(IsNull(M_OBJRS!DetailsR), "0", M_OBJRS!DetailsR)))
    M_OBJRS1!F_CEK = Trim(CStr(IIf(IsNull(M_OBJRS!F_CEK), "0", M_OBJRS!F_CEK)))
    M_OBJRS1!VisitKe = Trim(CStr(IIf(IsNull(M_OBJRS!VisitKe), "0", M_OBJRS!VisitKe)))
    M_OBJRS1!AddressToVisit = Trim(CStr(IIf(IsNull(M_OBJRS!AddressToVisit), "", M_OBJRS!AddressToVisit)))
    M_OBJRS1!Principle = Trim(CStr(IIf(IsNull(M_OBJRS!Principle), "0", M_OBJRS!Principle)))
    M_OBJRS1!amountwo = Trim(CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo)))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
    'Resume

End Sub
Private Sub Isi_Report_PTP_lunas()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Isi_REportErr

M_RPTCONN.Execute "Delete * from TrackingRptPayment"

M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient

    CMDSQL = "select mgm.custid, mgm.name, mgm.agent, ttlPTP, jmlBayar, mgm.ttlPTP-jmlbayar as sisaPay, mgm.tdbdatePTP,usertbl.spvcode from mgm inner join(select custid, sum(payment)as jmlBayar from tbllunas  group by custid )as a on mgm.custid = a.custid INNER JOIN usertbl on usertbl.userid=mgm.agent where tglstatus between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'AND  spvcode between '" + Combo3(0).Text + "' and '" + Combo3(1).Text + "' ORDER BY mgm.agent"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
M_OBJRS1.Open "Select * from TrackingRptPayment", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 1
While Not M_OBJRS.EOF
    M_OBJRS1.AddNew
    M_OBJRS1!CustId = CStr(IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId))
    M_OBJRS1!Name = CStr(IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name))
    M_OBJRS1!agent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
    M_OBJRS1!ttlptp = CStr(IIf(IsNull(M_OBJRS!ttlptp), 0, M_OBJRS!ttlptp))
    M_OBJRS1!jmlBayar = CStr(IIf(IsNull(M_OBJRS!jmlBayar), 0, M_OBJRS!jmlBayar))
    M_OBJRS1!SisaPay = CStr(IIf(IsNull(M_OBJRS!SisaPay), 0, M_OBJRS!SisaPay))
    M_OBJRS1!TglPTP = CStr(IIf(IsNull(M_OBJRS!TdbDatePTP), "2020-12-01", M_OBJRS!TdbDatePTP))
    M_OBJRS1!SPVCODE = CStr(IIf(IsNull(M_OBJRS!SPVCODE), "", M_OBJRS!SPVCODE))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

Isi_REportErr:
    MsgBox Err.Description
    'Resume
End Sub



Private Sub Isi_Agent_mgm()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr
Dim CMDSQL As String
Dim m_msgbox As Variant

M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"

M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    If Check1(2).Value <> 1 Then
        CMDSQL = "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' AND SPVCODE<>'SPV10' AND USERTYPE =1 ORDER BY USERID"
    Else
        CMDSQL = "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' AND USERTYPE =1 ORDER BY USERID"
    End If
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' AND USERTYPE =1 ORDER BY USERID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND USERTYPE =1 ORDER BY USERID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
M_OBJRS1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    M_OBJRS1.AddNew
    M_OBJRS1!TEAM = Trim(CStr(IIf(IsNull(M_OBJRS!TEAM), "", M_OBJRS!TEAM)))
    M_OBJRS1!TSRNAME = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    M_OBJRS1!TEAM = Trim(CStr(IIf(IsNull(M_OBJRS!SPVCODE), "", M_OBJRS!SPVCODE)))
    M_OBJRS1!AOC = Trim(CStr(IIf(IsNull(M_OBJRS!USERID), "", M_OBJRS!USERID)))
    M_OBJRS1!joindate = IIf(IsNull(M_OBJRS!joindate), "01/01/1900", M_OBJRS!joindate)
    M_OBJRS1!TARGET = IIf(IsNull(M_OBJRS!TARGET), 0, M_OBJRS!TARGET)
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description & "di procedure isi agent "
    'Resume
End Sub

Private Sub Isi_Agent_mgm_TARIK()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr
Dim CMDSQL As String
M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"

M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    CMDSQL = "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' AND USERTYPE =1 "
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' AND USERTYPE =1 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
M_OBJRS1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    M_OBJRS1.AddNew
    M_OBJRS1!TEAM = Trim(CStr(IIf(IsNull(M_OBJRS!TEAM), "", M_OBJRS!TEAM)))
    M_OBJRS1!TSRNAME = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    M_OBJRS1!TEAM = Trim(CStr(IIf(IsNull(M_OBJRS!SPVCODE), "", M_OBJRS!SPVCODE)))
    M_OBJRS1!AOC = Trim(CStr(IIf(IsNull(M_OBJRS!USERID), "", M_OBJRS!USERID)))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
    'Resume
End Sub

Private Sub Isi_Field_Collector()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingRptField"

M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' AND USERTYPE =2 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' AND USERTYPE =2 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        M_OBJRS.Open "Select * from usertbl where AKTIF = 0 AND USERTYPE =2", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
M_OBJRS1.Open "Select * from TrackingRptField", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    M_OBJRS1.AddNew
    M_OBJRS1!TEAM = Trim(CStr(IIf(IsNull(M_OBJRS!TEAM), "", M_OBJRS!TEAM)))
    M_OBJRS1!TSRNAME = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    M_OBJRS1!TEAM = Trim(CStr(IIf(IsNull(M_OBJRS!SPVCODE), "", M_OBJRS!SPVCODE)))
    M_OBJRS1!AOC = Trim(CStr(IIf(IsNull(M_OBJRS!USERID), "", M_OBJRS!USERID)))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
    'Resume
End Sub

Private Sub Tracking_PTP_Report()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim CMDSQL As String
'On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingPTP"
Set M_OBJRS = New ADODB.Recordset
Set M_OBJRS1 = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient
If Option1(1).Value = True Then

'CMDSQL = "select * from vwpay where f_cek like 'ptp%' and  tglstatus "
'CMDSQL = CMDSQL + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'CMDSQL = CMDSQL + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
'CMDSQL = CMDSQL + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') order by agent"
CMDSQL = "Select * from reportptp where f_cek like 'PTP%' and tglstatus  "
CMDSQL = CMDSQL + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
CMDSQL = CMDSQL + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
CMDSQL = CMDSQL + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') ORDER BY AGENT"
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then

'cmdsql = "select * from vwpay  where f_cek like 'ptp%' and tglstatus "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' ORDER BY AGENT  "

CMDSQL = "Select * from reportptp where f_cek like 'PTP%' and TglStatus  "
CMDSQL = CMDSQL + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
CMDSQL = CMDSQL + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' ORDER BY AGENT "

M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
M_OBJRS1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 1
While Not M_OBJRS.EOF
ProgressBar1.Value = M_OBJRS.Bookmark
    M_OBJRS1.AddNew
    M_OBJRS1!agent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
    M_OBJRS1!CustId = CStr(IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId))
    M_OBJRS1!Name = CStr(IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name))
    M_OBJRS1!TglPTP = CStr(IIf(IsNull(M_OBJRS!PromiseDate), "2020-12-30", M_OBJRS!PromiseDate))
    M_OBJRS1!ttlptp = CStr(IIf(IsNull(M_OBJRS!PromisePay), "0", M_OBJRS!PromisePay))
    'M_OBJRS1!BaseOn = CStr(IIf(IsNull(M_OBJRS!CmbBaseOn), "", M_OBJRS!CmbBaseOn))
    M_OBJRS1!Principle = CStr(IIf(IsNull(M_OBJRS!Principal), "0", M_OBJRS!Principal))
    M_OBJRS1!amountwo = CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub

'Isi_AgentErr:
 '   MsgBox Err.Description
    'Resume
End Sub
Private Sub hitung_JmlData_PerAgent_mgm()
Dim m_msgbox As Variant

'On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    JUMLAHVOL = Trim(CCur(IIf(IsNull(M_OBJRS!JMLVOL), "0", M_OBJRS!JMLVOL)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + ", JMLVOL= " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
  M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
JUMLAHVOL = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
 '       MsgBox Err.Description & "eror di hitung jumlah data peragent mgm"
    End If
    Resume Next
End Sub

Private Sub hitung_JmlData_PerAgent_mgm_TARIK()
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgmTARIK  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    JUMLAHVOL = Trim(CStr(IIf(IsNull(M_OBJRS!JMLVOL), "0", M_OBJRS!JMLVOL)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + ", JMLVOL= " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
JUMLAHVOL = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next
End Sub

Private Sub hitung_JmlData_FieldCollector()
Dim M_OBJRS As New ADODB.Recordset
Dim JUMLAH As String
Dim JUMLAHVOL As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "select FFC,count(FFC) as jml, sum(mgm.Amountwo) as JMLVOL from tblvisit INNER JOIN "
CMDSQL = CMDSQL + " mgm on tblVisit.custid=mgm.custid group by FFC "
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    JUMLAHVOL = Trim(CStr(IIf(IsNull(M_OBJRS!JMLVOL), "0", M_OBJRS!JMLVOL)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!FFC), "", M_OBJRS!FFC)))
    CMDSQL = "Update TrackingRptField set DATASIZE =" + JUMLAH + ", JMLVOL= " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next
End Sub
Private Sub Hitung_TrackingReportPerAgent_mgm()
Dim M_OBJRS As ADODB.Recordset
Dim M_OBJRS1 As ADODB.Recordset
Dim CMDSQL As String
Dim LAgent As String

On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    CMDSQL = "Select AGENT, kethslkerja, count(AGENT) as jumlah from mgm where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'group by AGENT, kethslkerja"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 1
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
'         WaitSecs (0.5)
        LAgent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        CMDSQL = CMDSQL + "[" + Trim(CStr(M_OBJRS!KETHSLKERJA)) + "]"
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(M_OBJRS!JUMLAH), 0, M_OBJRS!JUMLAH)) + ""
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(M_OBJRS!KETHSLKERJA) Then
        Else
            If M_OBJRS!KETHSLKERJA = "[]" Then
            Else
                If M_OBJRS!JUMLAH = 0 Then
                Else
                   
                    M_RPTCONN.Execute CMDSQL
                End If
            End If
        End If
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
'Resume
End Sub
Private Sub hitung_BatchCallInitilized_PerAgent_mgm()
Dim M_OBJRS As New ADODB.Recordset
'Dim JUMLAH As String
'Dim batch As String
'Dim CMDSQL As String
'Dim LAgent As String
Dim m_msgbox As Variant

On Error GoTo hitung_BatchCallInitilizeder
M_OBJRS.CursorLocation = adUseClient
CMDSQL = "Select agent,count(agent) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by agent order by  agent"
'CMDSQL = "Select userid,count(userid) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and left(RecSource,3) <> 'PRE' and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by userid"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "", M_OBJRS!Jml))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute CMDSQL
    End If
    M_OBJRS.MoveNext
Wend
    LAgent = Empty
    CMDSQL = Empty
    JUMLAH = Empty
    Set M_OBJRS = Nothing
    Exit Sub
hitung_BatchCallInitilizeder:
m_msgbox = MsgBox(Err.Description & "eror di hitung_batchcallinitilezed", vbRetryCancel, "Telegrandi")
If m_msgbox = vbRetry Then
    WaitSecs (3)
    Resume
End If
End Sub

Private Sub hitung_BatchCallInitilized_PerAgent_mgm_TARIK()
Dim M_OBJRS As New ADODB.Recordset
'Dim JUMLAH As String
'Dim batch As String
'Dim CMDSQL As String
'Dim LAgent As String
Dim m_msgbox As Variant

On Error GoTo hitung_BatchCallInitilizeder
M_OBJRS.CursorLocation = adUseClient
CMDSQL = "Select agent,count(agent) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and custid in(select custid from mgmTARIK where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by agent order by  agent"
'CMDSQL = "Select userid,count(userid) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and left(RecSource,3) <> 'PRE' and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by userid"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "", M_OBJRS!Jml))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute CMDSQL
    End If
    M_OBJRS.MoveNext
Wend
    LAgent = Empty
    CMDSQL = Empty
    JUMLAH = Empty
    Set M_OBJRS = Nothing
    Exit Sub
hitung_BatchCallInitilizeder:
m_msgbox = MsgBox(Err.Description, vbRetryCancel)
If m_msgbox = vbRetry Then
    WaitSecs (3)
    Resume
End If
End Sub

Private Sub hitung_BatchCallInitilized_PerAgent_Compare()
Dim M_OBJRS As New ADODB.Recordset
Dim m_msgbox As Variant

On Error GoTo hitung_BatchCallInitilizeder
M_OBJRS.CursorLocation = adUseClient
CMDSQL = "Select agent,count(agent) as jml from mgm_hst where tgl between '" + Format(TDBDate1(2).Value, "yyyy-mm-dd") & " " & DTimeLastCall(2).Value & "' and '" + Format(TDBDate1(3).Value, "yyyy-mm-dd") & " " & DTimeLastCall(3).Value & "'  and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by agent order by  agent"
'CMDSQL = "Select userid,count(userid) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and left(RecSource,3) <> 'PRE' and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by userid"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_OBJRS!Jml), "", M_OBJRS!Jml))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set [CALLSINITIATED_LM] ='" + JUMLAH + "' where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute CMDSQL
    End If
    M_OBJRS.MoveNext
   
Wend
Set M_OBJRS = Nothing
LAgent = Empty
CMDSQL = Empty
JUMLAH = Empty

Exit Sub
hitung_BatchCallInitilizeder:
m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
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
    RPT.Reset
End Sub


Private Sub header()
    listview1.ColumnHeaders.ADD 1, , "No", 4 * TXT
    listview1.ColumnHeaders.ADD 2, , "Report", 50 * TXT
End Sub

Private Sub Form_Load()
Dim listitem As listitem
Dim CMDSQL As String
Dim M_OBJRS As ADODB.Recordset
Set M_OBJRS = New ADODB.Recordset
DTimeLastCall(0).Value = "00:00"
DTimeLastCall(1).Value = "23:59"
DTimeLastCall(2).Value = "00:00"
DTimeLastCall(3).Value = "23:59"
M_OBJRS.CursorLocation = adUseClient
CmbCek.AddItem "Not Check"
CmbCek.AddItem "Accept"
CmbCek.AddItem "RETURN"

'If Check1(0).Value = 1 Then
'    cmdsql = "SELECT * FROM usertbl WHERE AKTIF = 0 AND USERID IN ("
'    cmdsql = cmdsql + "SELECT AGENT FROM mgm WHERE RECSOURCE LIKE '%CF%' AND RECSOURCE NOT LIKE '%CFR%') ORDER BY USERID"
    If Trim(UCase(MDIForm1.Text2)) = "ADMIN" Or Trim(UCase(MDIForm1.Text2)) = "SUPERVISOR" Then
    CMDSQL = "SELECT * FROM usertbl WHERE USERTYPE='1' order by USERID "
    Else
    CMDSQL = " select Userid from usertbl where team in (SELECT TEAM FROM usertbl WHERE userid = '" + Trim(MDIForm1.Text1.Text) + "'  )"
    End If
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_OBJRS.EOF
        Combo2(0).AddItem M_OBJRS!USERID
        Combo2(1).AddItem M_OBJRS!USERID
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
'End If

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "SELECT * FROM DATASOURCETBL where substring(kodeds,1,3) <> 'PRE' ORDER BY KODEDS", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJRS.EOF
    Combo1(0).AddItem M_OBJRS!KODEDS
    Combo1(1).AddItem M_OBJRS!KODEDS
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "SELECT * FROM usertbl where usertype='20' and spvcode like '%TL%' ORDER BY SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
M_OBJRS.Open "select *  from spvtbl order by SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 
While Not M_OBJRS.EOF
    Combo3(0).AddItem M_OBJRS!SPVCODE
    Combo3(1).AddItem M_OBJRS!SPVCODE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing

Call header

' report baru
'Set listitem = ListView1.ListItems.ADD(, , "1")
'    listitem.SubItems(1) = "CH Data Tracking Summary PerSPV Type A"
'Set listitem = ListView1.ListItems.ADD(, , "2")
'    listitem.SubItems(1) = "CH Data Tracking Summary PerDCR Name Type A"
Set listitem = listview1.ListItems.ADD(, , "1")
    listitem.SubItems(1) = "CH Data Tracking PerSPV Name Type B"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
Set listitem = listview1.ListItems.ADD(, , "2")
        listitem.SubItems(1) = "CH Data Tracking PerDCR Name Type B"
'Set listitem = ListView1.ListItems.ADD(, , "20")
'        listitem.SubItems(1) = "CH Data Tracking PerDCR Name Type B (Yg sudah ditarik)"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
Set listitem = listview1.ListItems.ADD(, , "3")
    listitem.SubItems(1) = "Status Call Data Tracking PerDCR Name Type B"
Set listitem = listview1.ListItems.ADD(, , "4")
    listitem.SubItems(1) = "Report Distribusi"
'Set listitem = ListView1.ListItems.ADD(, , "6")
'    listitem.SubItems(1) = "Report History Call"
'Set listitem = ListView1.ListItems.ADD(, , "7")
'    listitem.SubItems(1) = "Report History Call Group By CustID"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
Set listitem = listview1.ListItems.ADD(, , "5")
    listitem.SubItems(1) = "Promise To Pay Report"
'Set listitem = ListView1.ListItems.ADD(, , "9")
'    listitem.SubItems(1) = "Report Actual Payment"
Set listitem = listview1.ListItems.ADD(, , "6")
    listitem.SubItems(1) = "Report PTP Jatuh Tempo"
Set listitem = listview1.ListItems.ADD(, , "7")
    listitem.SubItems(1) = "Reserved PTP"
'Set listitem = ListView1.ListItems.ADD(, , "22")
'    listitem.SubItems(1) = "Report Deskcall Collector Collectivity"
''Set listitem = ListView1.ListItems.ADD(, , "11")
''    listitem.SubItems(1) = "Report History CH Sebelumnya"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
'Set listitem = ListView1.ListItems.ADD(, , "11")
'    listitem.SubItems(1) = "Report Form Visit"
'Set listitem = ListView1.ListItems.ADD(, , "12")
'    listitem.SubItems(1) = "Report Tracking Field Collector"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
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

'Penamabahan Report baru!
Set listitem = listview1.ListItems.ADD(, , "8")
    listitem.SubItems(1) = "Report POP BP1"
Set listitem = listview1.ListItems.ADD(, , "9")
    listitem.SubItems(1) = "Report POP BP2"
Set listitem = listview1.ListItems.ADD(, , "10")
    listitem.SubItems(1) = "Report POP BP3"
Set listitem = listview1.ListItems.ADD(, , "11")
    listitem.SubItems(1) = "Report BP1"
Set listitem = listview1.ListItems.ADD(, , "12")
    listitem.SubItems(1) = "Report BP2"
Set listitem = listview1.ListItems.ADD(, , "13")
    listitem.SubItems(1) = "Report BP3"
End Sub
Private Sub Form_Unload(Cancel As Integer)
M_OBJCONN.Close
Set M_OBJCONN = Nothing
M_OBJCONN.Open CMDSQLOPEN
End Sub
Private Sub ListView1_Click()
    Label2.Caption = listview1.SelectedItem.SubItems(1)
    Select Case listview1.SelectedItem.Index
    
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
Dim M_OBJRS As New ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_OBJRS.CursorLocation = adUseClient

    CMDSQL = "select agent, sum(Payment) as payment from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' group by agent"
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_OBJRS.RecordCount + 2
    While Not M_OBJRS.EOF
        ProgressBar1.Value = M_OBJRS.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_OBJRS!Payment), "0", M_OBJRS!Payment))
        LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
        CMDSQL = "Update TrackingRptPerPrgBatch set Payment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    JUMLAH = Empty
    LAgent = Empty
    CMDSQL = Empty
    Exit Sub
hitung_JmlDataer:
        If Err.Number = -2147217871 Then
            m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox Err.Description
        End If
End Sub

Public Sub hitung_jmlPTP()
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "select agent,count(custid) as jml from vwlunas where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND paydate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"

CMDSQL = "SELECT AGENT,sum(jumlah) as jml from (SELECT agent,count(custid) as jumlah FROM mgm"
'CMDSQL = CMDSQL + " WHERE F_CEK IN ('PTP','POP') and TGLSTATUS between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " WHERE left(F_CEK,3) LIKE '%PTP%' and TGLSTATUS between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and custid in (select custid from tblnegoptp where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "')"
CMDSQL = CMDSQL + " Group by agent,custid) a group by agent"

'CMDSQL = "select agent,sum(jumlah) as jml from (select mgm.agent,count(tblnegoptp.custid) as jumlah,tblnegoptp.promisedate FROM TBLnegoPTP,mgm Where mgm.CustId = tblnegoptp.CustId"
'CMDSQL = CMDSQL + " and promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
'CMDSQL = CMDSQL + " GROUP BY tblnegoptp.custid,promisedate,agent) a group by agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set PTP_BARU_LM =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next

End Sub


Public Sub hitung_jmlPOP_LM()
Dim m_msgbox As Variant
Dim bln As Integer
Dim TGL As Date
Dim mon As String
On Error GoTo hitung_JmlDataer
Dim CMDSQL As String

TGL = TDBDate1(0).Value
bln = DatePart("m", TGL) - 1
mon = "m"

M_OBJRS.CursorLocation = adUseClient
CMDSQL = "SELECT AGENT,sum(jumlah) as jml from (SELECT agent,count(custid) as jumlah FROM mgm"
CMDSQL = CMDSQL + " WHERE F_CEK='POP' and datepart(""" & mon & """,tglstatus)<='" & bln & "'"
CMDSQL = CMDSQL + " Group by agent,custid) a group by agent"

'CMDSQL = "SELECT AGENT,SUM(JUMLAH) AS JML FROM (SELECT agent,count(custid) as JUMLAH FROM tbllunas WHERE datepart(""" & mon & """,paydate)<='" & bln & "'"
'CMDSQL = CMDSQL + " Group by agent,custid) A GROUP BY AGENT"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set POP_LM =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next

End Sub
Public Sub PTP_withpayment()
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "select agent,count(custid) as jml from vwlunas where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND paydate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "select agent,sum(jumlah) as jml from (select agent,count(custid) as jumlah from tbllunas where custid in (select custid from tblnegoptp where promisedate "
'CMDSQL = CMDSQL + "between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "')"
'CMDSQL = CMDSQL + " and paydate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
'CMDSQL = CMDSQL + " GROUP BY agent,custid) a group by agent,jumlah"

CMDSQL = "SELECT AGENT,SUM(JUMLAH) AS JML FROM (SELECT agent,count(custid) as"
'CMDSQL = CMDSQL + " JUMLAH from mgm WHERE F_CEK in ('PTP','POP') and TGLSTATUS between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " JUMLAH from mgm WHERE F_CEK in ('PRE','POP','SP-','PTP-NE','PTP-PO','PTP-PA') and TGLSTATUS between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
CMDSQL = CMDSQL + " and custid in (select custid from tblnegoptp where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and custid in ("
CMDSQL = CMDSQL + " select custid from tbllunas WHERE Paydate Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'))"
CMDSQL = CMDSQL + " Group by agent,custid) A GROUP BY AGENT"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set PTP_BARU =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next

End Sub

Public Sub hitung_jmlPOP()
Dim m_msgbox As Variant
Dim bln As Integer
Dim TGL As Date
Dim mon As String
On Error GoTo hitung_JmlDataer
Dim CMDSQL As String

TGL = TDBDate1(0).Value
'bln = DatePart("m", TGL) - 1
'mon = "m"

M_OBJRS.CursorLocation = adUseClient
'CMDSQL = "SELECT AGENT,SUM(JUMLAH) AS JML FROM (SELECT agent,count(custid) as JUMLAH FROM tbllunas WHERE custid in "
'CMDSQL = CMDSQL + "(select custid from mgm WHERE F_CEK='POP' and datepart(""" & mon & """,tglstatus)<='" & bln & "') and"
'CMDSQL = CMDSQL + " Paydate Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'"
'CMDSQL = CMDSQL + " Group by agent,custid) A GROUP BY AGENT"

CMDSQL = "SELECT AGENT,SUM(JUMLAH) AS JML FROM (SELECT agent,count(custid) as"
CMDSQL = CMDSQL + " JUMLAH from mgm WHERE F_CEK='POP' or f_cek like 'ptp%'  and custid in ("
CMDSQL = CMDSQL + " select custid from tbllunas WHERE Paydate Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "')"
CMDSQL = CMDSQL + " Group by agent,custid) A GROUP BY AGENT"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set POP2 =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next

End Sub
Public Sub hitungcountpop()
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "select agent,sum(Payment) as jml from vwamount where paydate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and f_cek like '%pop%' group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set AmountPOP =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next
End Sub
Public Sub HITUNGHACCOUNTPTP()
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_OBJRS.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
'CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "select agent,count(CUSTID) as jml from vwPTP1 where promisedate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and f_cek like '%PTP%' and custid not in(select custid from vwlunas) group by Agent"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 2
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    JUMLAH = Trim(CStr(IIf(IsNull(M_OBJRS!Jml), "0", M_OBJRS!Jml)))
    LAgent = Trim(CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent)))
    CMDSQL = "Update TrackingRptPerPrgBatch set TTLPTP =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
JUMLAH = Empty
Exit Sub
hitung_JmlDataer:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    Resume Next
End Sub
Private Sub TrackingReservedPTP()
Dim M_OBJRS As New ADODB.Recordset
Dim M_OBJRS1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim CMDSQL As String
'Dim jenis As String
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingPTP"
Set M_OBJRS = New ADODB.Recordset
Set M_OBJRS1 = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS1.CursorLocation = adUseClient

If Option1(1).Value = True Then
CMDSQL = "select custid,sum(promisepay) as ReservedPTP, recsource,agent,promisedate, [name],AMOUNTWO,ttlPTP  from reportptp where inputdate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
CMDSQL = CMDSQL + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in  "
CMDSQL = CMDSQL + "(select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "')  group by custid, recsource,agent,[name], promisedate,AMOUNTWO,TTLPTP"

'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in "
'cmdsql = cmdsql + " (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') "
'cmdsql = cmdsql + " ORDER BY AGENT"
 M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
    
CMDSQL = "select custid,sum(promisepay) as  ReservedPTP, recsource,agent,promisedate, [name], AMOUNTWO, ttlPTP from reportptp  "
CMDSQL = CMDSQL + " where Inputdate between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' AND "
CMDSQL = CMDSQL + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND AGENT >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " group by custid, recsource,agent,name, promisedate,AMOUNTWO, ttlPTP "
    
'cmdsql = "Select * from reportptp where f_cek like '%PTP%' and promisedate  "
'cmdsql = cmdsql + " Between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and "
'cmdsql = cmdsql + " RECSOURCE Between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "' "
'cmdsql = cmdsql + " ORDER BY AGENT "
'
        M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
M_OBJRS1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_OBJRS.RecordCount + 1
While Not M_OBJRS.EOF
ProgressBar1.Value = M_OBJRS.Bookmark
    M_OBJRS1.AddNew
    M_OBJRS1!agent = CStr(IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent))
    M_OBJRS1!CustId = CStr(IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId))
    M_OBJRS1!Name = CStr(IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name))
    M_OBJRS1!TglPTP = CStr(IIf(IsNull(M_OBJRS!PromiseDate), "2020-12-30", M_OBJRS!PromiseDate))
    M_OBJRS1!sumreserved = CStr(IIf(IsNull(M_OBJRS!ReservedPTP), "0", M_OBJRS!ReservedPTP))
    M_OBJRS1!ttlptp = CStr(IIf(IsNull(M_OBJRS!ttlptp), "0", M_OBJRS!ttlptp))
    'M_OBJRS1!Tenor = CStr(IIf(IsNull(M_OBJRS!Tenor), "0", M_OBJRS!Tenor))
    M_OBJRS1!amountwo = CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo))
    M_OBJRS1.UPDATE
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
Set M_OBJRS1 = Nothing
Exit Sub
Isi_AgentErr:
MsgBox Err.Description
End Sub

Private Sub ISI_DATA_POP_BP1()

'@@Report POP-BP1 [24-11-09] -- POSTGREE -- Ritpill
Dim a As String
M_RPTCONN.Execute "Delete From rptPOP_BP"
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
CMDSQL = "select * from vwlunasmax"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

CMDSQL = " select custid,name,Amountwo,agent from MGM where agent not in ('LUNAS','PULLOUT') and "
If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
CMDSQL = CMDSQL + " and custid in("
CMDSQL = CMDSQL + " select custid from vwlunasmax where datediff('month',tglbayar,now())=1)"
'CMDSQL = CMDSQL + " select custid from vwlunasmax where datediff(month,tglbayar,getdate())=1)"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount > 0 Then
ProgressBar1.Max = M_OBJRS.RecordCount
End If
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    CMDSQL = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    CMDSQL = CMDSQL + " ("
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId) + "', "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name) + "', "
    CMDSQL = CMDSQL + " " + CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo)) + ", "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent) + "', "
    CMDSQL = CMDSQL + " 'POP BP1') "
    M_RPTCONN.Execute CMDSQL

M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
    
    
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
        CMDSQL = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        CMDSQL = CMDSQL + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
        CMDSQL = CMDSQL + " group by custid "
   
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_OBJRS.RecordCount > 0 Then
        ProgressBar1.Max = M_OBJRS.RecordCount
    End If
    While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set payment = " + CStr(M_OBJRS!Payment) + "  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from usertbl where  "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo3(0).Text + "' and userid <= '" + Combo3(1).Text + "' "
   End If
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
        While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set spvcode = '" + CStr(M_OBJRS!SPVCODE) + "'  where agent='" + Trim(M_OBJRS!USERID) + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "')"
    ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent > '" + Combo3(0).Text + "' and agent < '" + Combo3(1).Text + "')"
    End If
    
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
   While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set PaymentDate = '" + Format(M_OBJRS!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   

End Sub

Private Sub ISI_DATA_POP_BP2()

'@@Report POP-BP2 [24-11-09] -- POSTGREE -- Ritpill
Dim a As String
M_RPTCONN.Execute "Delete From rptPOP_BP"
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient


'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
CMDSQL = "select * from vwlunasmax"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient



CMDSQL = " select custid,name,Amountwo,agent from MGM where agent not in ('LUNAS','PULLOUT') and "
If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
CMDSQL = CMDSQL + " and custid in("
CMDSQL = CMDSQL + " select custid from vwlunasmax where datediff('month',tglbayar,now())=2)"
'CMDSQL = CMDSQL + " select custid from vwlunasmax where datediff(month,tglbayar,getdate())=1)"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount > 0 Then
ProgressBar1.Max = M_OBJRS.RecordCount
End If
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    CMDSQL = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    CMDSQL = CMDSQL + " ("
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId) + "', "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name) + "', "
    CMDSQL = CMDSQL + " " + CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo)) + ", "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent) + "', "
    CMDSQL = CMDSQL + " 'POP BP2') "
    M_RPTCONN.Execute CMDSQL

M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
    
    
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
        CMDSQL = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        CMDSQL = CMDSQL + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
        CMDSQL = CMDSQL + " group by custid "
   
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_OBJRS.RecordCount > 0 Then
        ProgressBar1.Max = M_OBJRS.RecordCount
    End If
    While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set payment = " + CStr(M_OBJRS!Payment) + "  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from usertbl where  "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo3(0).Text + "' and userid <= '" + Combo3(1).Text + "' "
   End If
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
        While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set spvcode = '" + CStr(M_OBJRS!SPVCODE) + "'  where agent='" + Trim(M_OBJRS!USERID) + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "')"
    ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent > '" + Combo3(0).Text + "' and agent < '" + Combo3(1).Text + "')"
    End If
    
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
   While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set PaymentDate = '" + Format(M_OBJRS!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing

End Sub

Private Sub ISI_DATA_POP_BP3()

'@@Report POP-BP3 [24-11-09] -- POSTGREE -- Ritpill
Dim a As String
M_RPTCONN.Execute "Delete From rptPOP_BP"
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
CMDSQL = "select * from vwlunasmax"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

CMDSQL = " select custid,name,Amountwo,agent from MGM where agent not in ('LUNAS','PULLOUT') and "
If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
CMDSQL = CMDSQL + " and custid in("
CMDSQL = CMDSQL + " select custid from vwlunasmax where datediff('month',tglbayar,now())>=3)"
'CMDSQL = CMDSQL + " select custid from vwlunasmax where datediff(month,tglbayar,getdate())=1)"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount > 0 Then
ProgressBar1.Max = M_OBJRS.RecordCount
End If
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    CMDSQL = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    CMDSQL = CMDSQL + " ("
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId) + "', "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name) + "', "
    CMDSQL = CMDSQL + " " + CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo)) + ", "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent) + "', "
    CMDSQL = CMDSQL + " 'POP BP3') "
    M_RPTCONN.Execute CMDSQL

M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
    
    
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
        CMDSQL = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        CMDSQL = CMDSQL + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
        CMDSQL = CMDSQL + " group by custid "
   
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_OBJRS.RecordCount > 0 Then
        ProgressBar1.Max = M_OBJRS.RecordCount
    End If
    While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set payment = " + CStr(M_OBJRS!Payment) + "  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from usertbl where  "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo3(0).Text + "' and userid <= '" + Combo3(1).Text + "' "
   End If
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
        While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set spvcode = '" + CStr(M_OBJRS!SPVCODE) + "'  where agent='" + Trim(M_OBJRS!USERID) + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "')"
    ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent > '" + Combo3(0).Text + "' and agent < '" + Combo3(1).Text + "')"
    End If
    
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
   While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set PaymentDate = '" + Format(M_OBJRS!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing


End Sub

Private Sub ISI_DATA_BP1()

'@@Report BP1 [24-11-09] -- POSTGREE -- Ritpill
Dim a As String
M_RPTCONN.Execute "Delete From rptPOP_BP"
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient


'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
CMDSQL = "select * from vwptp1"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient


CMDSQL = " select custid,name,Amountwo,agent from MGM where "
If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
CMDSQL = CMDSQL + " and custid in("
CMDSQL = CMDSQL + " select custid from vwptp1 where datediff('month',promisedate,now())=1 and custid not in (select distinct custid from tbllunas))"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount > 0 Then
ProgressBar1.Max = M_OBJRS.RecordCount
End If
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    CMDSQL = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    CMDSQL = CMDSQL + " ("
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId) + "', "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name) + "', "
    CMDSQL = CMDSQL + " " + CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo)) + ", "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent) + "', "
    CMDSQL = CMDSQL + " 'BP1') "
    M_RPTCONN.Execute CMDSQL

M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
    
    
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
        CMDSQL = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        CMDSQL = CMDSQL + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
        CMDSQL = CMDSQL + " group by custid "
   
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_OBJRS.RecordCount > 0 Then
        ProgressBar1.Max = M_OBJRS.RecordCount
    End If
    While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set payment = " + CStr(M_OBJRS!Payment) + "  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from usertbl where  "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo3(0).Text + "' and userid <= '" + Combo3(1).Text + "' "
   End If
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
        While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set spvcode = '" + CStr(M_OBJRS!SPVCODE) + "'  where agent='" + Trim(M_OBJRS!USERID) + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "')"
    ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent > '" + Combo3(0).Text + "' and agent < '" + Combo3(1).Text + "')"
    End If
    
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
   While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set PaymentDate = '" + Format(M_OBJRS!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   

End Sub

Private Sub ISI_DATA_BP2()

'@@Report BP2 [24-11-09] -- POSTGREE -- Ritpill
Dim a As String
M_RPTCONN.Execute "Delete From rptPOP_BP"
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
CMDSQL = "select * from vwptp1"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

CMDSQL = " select custid,name,Amountwo,agent from MGM where "
If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
CMDSQL = CMDSQL + " and custid in("
CMDSQL = CMDSQL + " select custid from vwptp1 where datediff('month',promisedate,now())=2 and custid not in (select distinct custid from tbllunas))"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount > 0 Then
ProgressBar1.Max = M_OBJRS.RecordCount
End If
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    CMDSQL = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    CMDSQL = CMDSQL + " ("
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId) + "', "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name) + "', "
    CMDSQL = CMDSQL + " " + CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo)) + ", "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent) + "', "
    CMDSQL = CMDSQL + " 'BP2') "
    M_RPTCONN.Execute CMDSQL

M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
    
    
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
        CMDSQL = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        CMDSQL = CMDSQL + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
        CMDSQL = CMDSQL + " group by custid "
   
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_OBJRS.RecordCount > 0 Then
        ProgressBar1.Max = M_OBJRS.RecordCount
    End If
    While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set payment = " + CStr(M_OBJRS!Payment) + "  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from usertbl where  "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo3(0).Text + "' and userid <= '" + Combo3(1).Text + "' "
   End If
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
        While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set spvcode = '" + CStr(M_OBJRS!SPVCODE) + "'  where agent='" + Trim(M_OBJRS!USERID) + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "')"
    ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent > '" + Combo3(0).Text + "' and agent < '" + Combo3(1).Text + "')"
    End If
    
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
   While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set PaymentDate = '" + Format(M_OBJRS!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   

End Sub

Private Sub ISI_DATA_BP3()

'@@Report BP3 [24-11-09] -- POSTGREE -- Ritpill
Dim a As String
M_RPTCONN.Execute "Delete From rptPOP_BP"
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

'@@ Error Handler, jika tabel vwlunasmx kosong! [25-11-09]
CMDSQL = "select * from vwptp1"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount = 0 Then
 a = MsgBox("Data tidak ada!", vbOKOnly + vbInformation, "informasi")
 Exit Sub
End If
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient

CMDSQL = " select custid,name,Amountwo,agent from MGM where "
If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
CMDSQL = CMDSQL + " and custid in("
CMDSQL = CMDSQL + " select custid from vwptp1 where datediff('month',promisedate,now())>=3 and custid not in (select distinct custid from tbllunas))"
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If M_OBJRS.RecordCount > 0 Then
ProgressBar1.Max = M_OBJRS.RecordCount
End If
While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
    CMDSQL = "INSERT INTO RptPOP_BP(custid,name,Amountwo,agent,category)VALUES "
    CMDSQL = CMDSQL + " ("
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!CustId), "", M_OBJRS!CustId) + "', "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!Name), "", M_OBJRS!Name) + "', "
    CMDSQL = CMDSQL + " " + CStr(IIf(IsNull(M_OBJRS!amountwo), "0", M_OBJRS!amountwo)) + ", "
    CMDSQL = CMDSQL + " '" + IIf(IsNull(M_OBJRS!agent), "", M_OBJRS!agent) + "', "
    CMDSQL = CMDSQL + " 'BP3') "
    M_RPTCONN.Execute CMDSQL

M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
    
    
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
        CMDSQL = " select custid,sum(payment)as payment from reportPOP where paydate > tglsource and custid in (select custid from mgm where agent not in"
        CMDSQL = CMDSQL + " ('LUNAS','PULLOUT')) AND "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo2(0).Text + "' and agent <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " agent >= '" + Combo3(0).Text + "' and agent <= '" + Combo3(1).Text + "' "
   End If
        CMDSQL = CMDSQL + " group by custid "
   
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   If M_OBJRS.RecordCount > 0 Then
        ProgressBar1.Max = M_OBJRS.RecordCount
    End If
    While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set payment = " + CStr(M_OBJRS!Payment) + "  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from usertbl where  "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' "
   ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " userid >= '" + Combo3(0).Text + "' and userid <= '" + Combo3(1).Text + "' "
   End If
   
   M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
  If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
        While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set spvcode = '" + CStr(M_OBJRS!SPVCODE) + "'  where agent='" + Trim(M_OBJRS!USERID) + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   
   
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   CMDSQL = " select * from vwlunasmax where custid in "
    If Option1(0).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent >= '" + Combo1(0).Text + "' and agent <= '" + Combo1(1).Text + "')"
    ElseIf Option1(1).Value = True Then
        CMDSQL = CMDSQL + " (select custid from mgm where agent > '" + Combo3(0).Text + "' and agent < '" + Combo3(1).Text + "')"
    End If
    
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
   If M_OBJRS.RecordCount > 0 Then
    ProgressBar1.Max = M_OBJRS.RecordCount
    End If
   While Not M_OBJRS.EOF
    ProgressBar1.Value = M_OBJRS.Bookmark
        CMDSQL = " update RptPOP_BP set PaymentDate = '" + Format(M_OBJRS!tglbayar, "yyyy-mm-dd") + "'  where custid='" + M_OBJRS!CustId + "'"
        M_RPTCONN.Execute CMDSQL
   M_OBJRS.MoveNext
   Wend
   Set M_OBJRS = Nothing
   

End Sub



