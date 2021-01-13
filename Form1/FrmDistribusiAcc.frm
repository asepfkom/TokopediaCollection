VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FrmDistribusiAcc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manage distribusi account"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15660
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   15660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "Monitoring"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd_insert_antrian 
      BackColor       =   &H0000FF00&
      Caption         =   "&Proses..."
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   6210
      Width           =   1395
   End
   Begin VB.CommandButton CmdFilterExcel 
      Caption         =   "&Filter dari Excel"
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
      Left            =   2700
      TabIndex        =   54
      Top             =   0
      Width           =   1995
   End
   Begin VB.ComboBox CmbStatusCollBersama 
      Height          =   315
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   360
      Width           =   1635
   End
   Begin VB.ComboBox CmbAgentCollBersama 
      Height          =   315
      Left            =   7980
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   360
      Width           =   1395
   End
   Begin VB.CommandButton CmdFormClaimAccount 
      BackColor       =   &H0000C0C0&
      Caption         =   "Form Claim Account..."
      Height          =   435
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6120
      Width           =   2355
   End
   Begin VB.ComboBox CmbStatusAcc 
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   360
      Width           =   1635
   End
   Begin VB.ComboBox CmbAgent 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   7080
      Width           =   2235
   End
   Begin VB.CommandButton CmdBukaAccount 
      Caption         =   "Buka account terkunci..."
      Height          =   435
      Left            =   10680
      TabIndex        =   36
      Top             =   6120
      Width           =   2355
   End
   Begin VB.CommandButton CmdKembalikanAgent 
      Caption         =   "Kembalikan Ke Agent lama..."
      Height          =   435
      Left            =   8280
      TabIndex        =   37
      Top             =   6120
      Width           =   2355
   End
   Begin VB.ComboBox CmbFilterAcc 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   360
      Width           =   1995
   End
   Begin VB.TextBox TxtJmlhAcc 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   27
      Text            =   "0"
      Top             =   4740
      Width           =   915
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   14520
      TabIndex        =   25
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton CmdCari 
      Caption         =   "&Cari"
      Height          =   315
      Left            =   14520
      TabIndex        =   24
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox TxtCariNama 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12660
      TabIndex        =   23
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox TxtCariCustid 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12660
      TabIndex        =   21
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton CmdUncekAll 
      Caption         =   "UnCek all"
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton CmdCekAllAcc 
      Caption         =   "Cek all"
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   60
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   2580
      TabIndex        =   17
      Top             =   4800
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox TxtJmlhAgent 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1620
      TabIndex        =   16
      Text            =   "0"
      Top             =   9900
      Width           =   915
   End
   Begin VB.CommandButton CmdProses 
      BackColor       =   &H0000FF00&
      Caption         =   "&Approve..."
      Height          =   375
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6210
      Width           =   1515
   End
   Begin VB.CommandButton CmdLihatListAgent 
      Caption         =   "Lihat list agent..."
      Height          =   435
      Left            =   10740
      TabIndex        =   3
      Top             =   5160
      Width           =   1755
   End
   Begin VB.TextBox TxtAgent 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   5220
      Width           =   7035
   End
   Begin MSComctlLib.ListView LvAcc 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   7011
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
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LvAgent 
      Height          =   2415
      Left            =   60
      TabIndex        =   6
      Top             =   7440
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   4260
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
      NumItems        =   0
   End
   Begin TDBDate6Ctl.TDBDate TxtTglAwal 
      Height          =   315
      Left            =   1260
      TabIndex        =   28
      Top             =   5640
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   556
      Calendar        =   "FrmDistribusiAcc.frx":0000
      Caption         =   "FrmDistribusiAcc.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":0184
      Keys            =   "FrmDistribusiAcc.frx":01A2
      Spin            =   "FrmDistribusiAcc.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
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
      Text            =   "____-__-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime TxtWaktuAwal 
      Height          =   315
      Left            =   2775
      TabIndex        =   29
      Top             =   5640
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmDistribusiAcc.frx":0228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDistribusiAcc.frx":0294
      Spin            =   "FrmDistribusiAcc.frx":02E4
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
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
   Begin TDBDate6Ctl.TDBDate TxtTglAkhir 
      Height          =   315
      Left            =   4800
      TabIndex        =   30
      Top             =   5640
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   556
      Calendar        =   "FrmDistribusiAcc.frx":030C
      Caption         =   "FrmDistribusiAcc.frx":0424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":0490
      Keys            =   "FrmDistribusiAcc.frx":04AE
      Spin            =   "FrmDistribusiAcc.frx":050C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
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
      Text            =   "____-__-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime TxtWaktuAkhir 
      Height          =   315
      Left            =   6315
      TabIndex        =   31
      Top             =   5640
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmDistribusiAcc.frx":0534
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDistribusiAcc.frx":05A0
      Spin            =   "FrmDistribusiAcc.frx":05F0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
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
   Begin VB.CommandButton CmdHapusAgent 
      Caption         =   "&Hapus Agent"
      Height          =   375
      Left            =   13500
      TabIndex        =   14
      Top             =   7800
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   13500
      TabIndex        =   15
      Top             =   7320
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton CmdCekAllAgent 
      Caption         =   "Cek All"
      Height          =   375
      Left            =   13500
      TabIndex        =   42
      Top             =   8400
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton CmdUncekallAgent 
      Caption         =   "UnCek All"
      Height          =   375
      Left            =   13500
      TabIndex        =   43
      Top             =   8820
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Filter dari Kriteria"
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
      Left            =   4800
      TabIndex        =   78
      Top             =   0
      Width           =   1995
   End
   Begin TDBDate6Ctl.TDBDate TxtTglExpired 
      Height          =   315
      Left            =   2100
      TabIndex        =   92
      Top             =   6000
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   556
      Calendar        =   "FrmDistribusiAcc.frx":0618
      Caption         =   "FrmDistribusiAcc.frx":0730
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":079C
      Keys            =   "FrmDistribusiAcc.frx":07BA
      Spin            =   "FrmDistribusiAcc.frx":0818
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
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
      Text            =   "____-__-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   5130
      TabIndex        =   56
      Top             =   330
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox Check_decease 
         Caption         =   "Include Account Decease [ 835 ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   7200
         Width           =   3375
      End
      Begin VB.ListBox list_batch 
         Height          =   1035
         ItemData        =   "FrmDistribusiAcc.frx":0840
         Left            =   1080
         List            =   "FrmDistribusiAcc.frx":0847
         MultiSelect     =   2  'Extended
         TabIndex        =   90
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
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
         Left            =   3360
         TabIndex        =   79
         Top             =   7080
         Width           =   1335
      End
      Begin VB.ComboBox cb_batch 
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
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   360
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Filter"
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
         Left            =   4680
         TabIndex        =   73
         Top             =   7080
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "WO DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   57
         Top             =   3690
         Width           =   2895
         Begin VB.CheckBox Check2 
            Caption         =   "WO DATE"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1095
         End
         Begin VB.ListBox List1 
            Enabled         =   0   'False
            Height          =   2400
            ItemData        =   "FrmDistribusiAcc.frx":0857
            Left            =   360
            List            =   "FrmDistribusiAcc.frx":088B
            MultiSelect     =   2  'Extended
            TabIndex        =   58
            Top             =   750
            Width           =   1455
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1575
         Left            =   120
         TabIndex        =   60
         Top             =   2040
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "STATUS"
         Begin VB.CheckBox Check1 
            Caption         =   "UN-Uncontacted"
            Height          =   255
            Index           =   7
            Left            =   150
            TabIndex        =   72
            Top             =   2010
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "POP-Progress Of Payment"
            Height          =   255
            Index           =   6
            Left            =   3030
            TabIndex        =   71
            Top             =   510
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "BP-Broken Promise"
            Height          =   255
            Index           =   5
            Left            =   3030
            TabIndex        =   70
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "PTP-Promise To Pay"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Top             =   2070
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "PR-PROSPECT"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VL-VALID"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "RP-Refuse Payment"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   66
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SP-Settled Payment"
            Height          =   255
            Index           =   8
            Left            =   3030
            TabIndex        =   65
            Top             =   750
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "BL-Data Blank"
            Height          =   255
            Index           =   9
            Left            =   150
            TabIndex        =   64
            Top             =   2010
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "OS-On Process"
            Height          =   255
            Index           =   10
            Left            =   3030
            TabIndex        =   63
            Top             =   990
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "ON-On Nego"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   62
            Top             =   990
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SK-SKIP"
            Height          =   255
            Index           =   11
            Left            =   3030
            TabIndex        =   61
            Top             =   1200
            Width           =   2535
         End
      End
      Begin MSComCtl2.DTPicker tgl_lpd 
         Height          =   375
         Left            =   1080
         TabIndex        =   76
         Top             =   1560
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "MMMM-yyyy"
         Format          =   69468163
         CurrentDate     =   41610
      End
      Begin MSComCtl2.DTPicker tgl_lpd2 
         Height          =   375
         Left            =   3840
         TabIndex        =   81
         Top             =   1560
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "MMMM-yyyy"
         Format          =   69468163
         CurrentDate     =   41610
      End
      Begin VB.Frame Frame3 
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   3120
         TabIndex        =   80
         Top             =   3690
         Width           =   2895
         Begin VB.Frame Frame4 
            Enabled         =   0   'False
            Height          =   2295
            Left            =   240
            TabIndex        =   83
            Top             =   600
            Width           =   2415
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "5.000.000 - 10.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   85
               Top             =   600
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "10.000.000 - 30.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   86
               Top             =   960
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "30.000.000 - 60.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   87
               Top             =   1320
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "60.000.000 - 90.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   88
               Top             =   1680
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "90.000.000 - 120.000.000"
            End
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Current Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   89
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "LPD"
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
         Left            =   240
         TabIndex        =   77
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.Label Label25 
      Caption         =   "AKSESALL MENUNGGU APPROVAL"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12840
      TabIndex        =   95
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label24 
      Caption         =   "Exp Date :"
      Height          =   225
      Left            =   1260
      TabIndex        =   91
      Top             =   6060
      Width           =   975
   End
   Begin VB.Label lbl_profile 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Profile >>>"
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
      Left            =   13320
      TabIndex        =   55
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label20 
      Caption         =   "Status:"
      Height          =   195
      Left            =   9360
      TabIndex        =   52
      Top             =   420
      Width           =   555
   End
   Begin VB.Label Label19 
      Caption         =   "Agent AWAL:"
      Height          =   195
      Left            =   6960
      TabIndex        =   50
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Filter Account Sedang di Collect Bersama:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   49
      Top             =   60
      Width           =   4215
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label LblWaktuServer 
      BackColor       =   &H000080FF&
      Caption         =   "<Waktu Server>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9540
      TabIndex        =   48
      Top             =   5640
      Width           =   2355
   End
   Begin VB.Label Label18 
      BackColor       =   &H000040C0&
      Caption         =   "Waktu Server Saat ini:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   47
      Top             =   5640
      Width           =   2355
   End
   Begin VB.Label Label17 
      Caption         =   "Status Acc:"
      Height          =   195
      Left            =   4380
      TabIndex        =   45
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label16 
      Caption         =   "Agent:"
      Height          =   195
      Left            =   1260
      TabIndex        =   44
      Top             =   420
      Width           =   555
   End
   Begin VB.Label Label15 
      Caption         =   "AND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3900
      TabIndex        =   40
      Top             =   420
      Width           =   435
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3300
      X2              =   3300
      Y1              =   7080
      Y2              =   7440
   End
   Begin VB.Label Label14 
      Caption         =   "Filter Agent:"
      Height          =   195
      Left            =   60
      TabIndex        =   38
      Top             =   7140
      Width           =   915
   End
   Begin VB.Label Label13 
      Caption         =   "Filter Account:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   34
      Top             =   60
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6900
      X2              =   6900
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label Label12 
      Caption         =   "Waktu Awal:"
      Height          =   195
      Left            =   210
      TabIndex        =   33
      Top             =   5700
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Waktu Akhir:"
      Height          =   195
      Left            =   3780
      TabIndex        =   32
      Top             =   5700
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Jumlah Account:"
      Height          =   195
      Left            =   180
      TabIndex        =   26
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Cari Nama:"
      Height          =   195
      Left            =   11700
      TabIndex        =   22
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label8 
      Caption         =   "Cari Custid:"
      Height          =   195
      Left            =   11700
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LblStatusAcc 
      Caption         =   "<none>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   13
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Status Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10980
      TabIndex        =   12
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label LblNama 
      Caption         =   "<none>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Nama:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7500
      TabIndex        =   10
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label LblCustid 
      Caption         =   "<none>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Custid terpilih:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3420
      TabIndex        =   8
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Jumlah Data Agent:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   9960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   $"FrmDistribusiAcc.frx":08EF
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   6720
      Width           =   15375
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   15420
      Y1              =   6660
      Y2              =   6660
   End
   Begin VB.Label Label2 
      Caption         =   "Agent yang boleh mengakses account di atas:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   5220
      Width           =   3735
   End
End
Attribute VB_Name = "FrmDistribusiAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub new_kdprofile()
    Dim M_Objrs As ADODB.Recordset
    Dim index_profile As Integer
    Dim tglprofile As String
    
    ' ------------ KODE PROFILE 21 MEI 2013 ------------------
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open "SELECT now()", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    tglprofile = Format(M_Objrs(0), "yyyymmdd")
    
    If M_Objrs.state = 1 Then M_Objrs.Close
    M_Objrs.Open "SELECT * FROM tbl_profile_aksesall"
    If M_Objrs.RecordCount > 0 Then
        index_profile = M_Objrs.RecordCount
        lbl_profile.Caption = tglprofile & Right("0000" & index_profile + 1, 4)
    Else
        lbl_profile.Caption = tglprofile & "0001"
    End If
    ' ---------------------------------------------------------
    
    Set M_Objrs = Nothing
End Sub

Private Sub HeaderAccount()
    LvAcc.ColumnHeaders.ADD 1, , "Custid", 2000
    LvAcc.ColumnHeaders.ADD 2, , "Nama Costumer", 3000
    LvAcc.ColumnHeaders.ADD 3, , "Status Account", 3000
    LvAcc.ColumnHeaders.ADD 4, , "Agent Saat ini", 1500
    LvAcc.ColumnHeaders.ADD 5, , "Agent Terdahulu", 1500
    LvAcc.ColumnHeaders.ADD 6, , "Akses Saat ini", 1500
    LvAcc.ColumnHeaders.ADD 7, , "Waktu Akses Saat ini", 1500
End Sub

Private Sub IsiAccount()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    Dim M_WHERE As String
    
    M_WHERE = ""
    
    cmdsql = "select * from mgm  "
    
    If TxtCariCustid.text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where custid like '%" + CStr(TxtCariCustid.text) + "%' "
        Else
            M_WHERE = M_WHERE & " and custid like '%" + CStr(TxtCariCustid.text) + "%' "
        End If
    End If
    
    If TxtCariNama.text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where name like '%" + CStr(TxtCariNama.text) + "%' "
        Else
            M_WHERE = M_WHERE & " and name like '%" + CStr(TxtCariNama.text) + "%' "
        End If
    End If
       
    If M_WHERE <> "" Then
        M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS','CLAIM')   "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
    Else
        M_WHERE = " where agent not in ('COMPLAIN','LUNAS','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
    End If
       
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql & M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.clear
    TxtJmlhAcc.text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set listItem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            listItem.SubItems(1) = M_Objrs("name")
            listItem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            listItem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            listItem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            listItem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            listItem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                listItem.ForeColor = vbRed
                listItem.ListSubItems(1).ForeColor = vbRed
                listItem.ListSubItems(2).ForeColor = vbRed
                listItem.ListSubItems(3).ForeColor = vbRed
                listItem.ListSubItems(4).ForeColor = vbRed
                listItem.ListSubItems(5).ForeColor = vbRed
                listItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                listItem.ForeColor = vbBlue
                listItem.ListSubItems(1).ForeColor = vbBlue
                listItem.ListSubItems(2).ForeColor = vbBlue
                listItem.ListSubItems(3).ForeColor = vbBlue
                listItem.ListSubItems(4).ForeColor = vbBlue
                listItem.ListSubItems(5).ForeColor = vbBlue
                listItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub HeaderAgent()
    LvAgent.ColumnHeaders.ADD 1, , "ID", 500
    LvAgent.ColumnHeaders.ADD 2, , "AGENT", 1000
    LvAgent.ColumnHeaders.ADD 3, , "CUSTID", 2000
    LvAgent.ColumnHeaders.ADD 4, , "WAKTU AWAL", 2000
    LvAgent.ColumnHeaders.ADD 5, , "WAKTU AKHIR", 2000
    LvAgent.ColumnHeaders.ADD 6, , "LOG DISTRIBUSI", 1500
    LvAgent.ColumnHeaders.ADD 7, , "WAKTU DISTRIBUSI", 2000
    LvAgent.ColumnHeaders.ADD 8, , "KODE PROFILE", 2000
End Sub




Private Sub Check2_Click()
    If Check2.Value = 1 Then
        List1.Enabled = True
    Else
        List1.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Frame4.Enabled = True
    Else
        Frame4.Enabled = False
    End If
End Sub

Private Sub CmbAgent_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    Dim GroupingTL_2 As String
    
    GroupingTL_2 = ""
    
    '@@19022013 Tambahan ini buat grouping TL
    If UCase(Mid(CmbAgent.text, 1, 2)) = "TL" Then
        GroupingTL_2 = " agent in (select userid from usertbl where spvcode in ("
        GroupingTL_2 = GroupingTL_2 & " select spvcode from usertbl where userid='"
        GroupingTL_2 = GroupingTL_2 & CmbAgent.text + "')) "
    Else
        GroupingTL_2 = " agent='"
        GroupingTL_2 = GroupingTL_2 + CmbAgent.text + "' "
    End If
    
    If CmbAgent.text <> "ALL" Then
        'Cmdsql = "select * from tbl_distribusi_account where " & GroupingTL_2
        'Cmdsql = Cmdsql & CmbAgent.Text & "' order by waktu_awal asc "
    Else
        cmdsql = "select * from tbl_distribusi_account "
        cmdsql = cmdsql & " order by agent,waktu_awal asc "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlhAgent.text = M_Objrs.RecordCount
    LvAgent.ListItems.clear
    
    If M_Objrs.RecordCount = 0 Then
        
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    While Not M_Objrs.EOF
        Set listItem = LvAgent.ListItems.ADD(, , M_Objrs("id"))
            listItem.SubItems(1) = M_Objrs("agent")
            listItem.SubItems(2) = M_Objrs("custid")
            listItem.SubItems(3) = Format(M_Objrs("waktu_awal"), "yyyy-mm-dd hh:nn:ss")
            listItem.SubItems(4) = Format(M_Objrs("waktu_akhir"), "yyyy-mm-dd hh:nn:ss")
            listItem.SubItems(5) = M_Objrs("log_distribusi")
            listItem.SubItems(6) = Format(M_Objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
            listItem.SubItems(7) = Format(M_Objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
       M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub CmbFilterAcc_Click()
    'Call CariFilter
    'Eventnya diambil berdasarkan
End Sub



Private Sub CmbStatusAcc_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    Dim Bulan, Tahun, Tanggal As String
    Dim M_WHERE As String
    Dim GroupingTL As String
    
    If CmbFilterAcc.text = "" Then
        'MsgBox "Pilih terlebih dahulu agentnya!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

    M_WHERE = ""
    GroupingTL = ""
    
    cmdsql = "select * from mgm "
    
    If CmbFilterAcc.text = "ALL" Then
        
        'Ini jika agent=ALL dan status account=ALL
        If CmbStatusAcc.text = "ALL" Then
            'CUKUP DEH NGGA USAH PAKE SCRIPT---------
        
        '@@15022013 Ini jika filter agentnya=ALL tetapi status accountnya <> ALL
        ElseIf CmbStatusAcc.text <> "ALL" Then
            If CmbStatusAcc.text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusAcc.text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusAcc.text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusAcc.text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusAcc.text <> "LPD 1" And _
                   CmbStatusAcc.text <> "LPD 2" And _
                   CmbStatusAcc.text <> "LPD 3" And _
                   CmbStatusAcc.text <> "LPD 3<" Then
                   
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusAcc.text & "%' "
                Else
                    M_WHERE = " and f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusAcc.text & "%' "
                End If
            End If
        End If
        
    ElseIf CmbFilterAcc.text <> "ALL" Then
    
        '@@19022013 Tambahan ini buat grouping TL
        If UCase(Mid(CmbFilterAcc.text, 1, 2)) = "TL" Then
            GroupingTL = " agent in (select userid from usertbl where spvcode in ("
            GroupingTL = GroupingTL & " select spvcode from usertbl where userid='"
            GroupingTL = GroupingTL & CmbFilterAcc.text + "')) "
        Else
            GroupingTL = " agent='"
            GroupingTL = GroupingTL + CmbFilterAcc.text + "' "
        End If
                
        'Ini jika agent <>ALL dan status account=ALL
        If CmbStatusAcc.text = "ALL" Then
            If M_WHERE = "" Then
                'M_WHERE = " where agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " where " & GroupingTL
            Else
                'M_WHERE = " and agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " and " & GroupingTL
            End If
            
        'Ini jika agent=<>ALL dan status account <> ALL
        ElseIf CmbStatusAcc.text <> "ALL" Then
            If CmbStatusAcc.text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusAcc.text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and  " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            End If
            
            If CmbStatusAcc.text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusAcc.text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusAcc.text <> "LPD 1" And _
                   CmbStatusAcc.text <> "LPD 2" And _
                   CmbStatusAcc.text <> "LPD 3" And _
                   CmbStatusAcc.text <> "LPD 3<" Then
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusAcc.text & "%' and " & GroupingTL
                Else
                    M_WHERE = " and f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusAcc.text & "%' and " & GroupingTL
                End If
            End If
        End If
            
    End If
    
    If M_WHERE = "" Then
        M_WHERE = " where agent not in ('LUNAS','COMPLAIN','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
        M_WHERE = M_WHERE & " order by name asc "
    Else
        M_WHERE = M_WHERE & " and agent not in ('LUNAS','COMPLAIN','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
        M_WHERE = M_WHERE & " order by name asc "
    End If
    
    
    
    DoEvents
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql + M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.clear
    TxtJmlhAcc.text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set listItem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            listItem.SubItems(1) = M_Objrs("name")
            listItem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            listItem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            listItem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            listItem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            listItem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                listItem.ForeColor = vbRed
                listItem.ListSubItems(1).ForeColor = vbRed
                listItem.ListSubItems(2).ForeColor = vbRed
                listItem.ListSubItems(3).ForeColor = vbRed
                listItem.ListSubItems(4).ForeColor = vbRed
                listItem.ListSubItems(5).ForeColor = vbRed
                listItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                listItem.ForeColor = vbBlue
                listItem.ListSubItems(1).ForeColor = vbBlue
                listItem.ListSubItems(2).ForeColor = vbBlue
                listItem.ListSubItems(3).ForeColor = vbBlue
                listItem.ListSubItems(4).ForeColor = vbBlue
                listItem.ListSubItems(5).ForeColor = vbBlue
                listItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    
End Sub



Private Sub CmbStatusCollBersama_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    Dim Bulan, Tahun, Tanggal As String
    Dim M_WHERE As String
    Dim GroupingTL As String
    
    If CmbStatusCollBersama.text = "" Then
        MsgBox "Mohon maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbAgentCollBersama.text = "" Then
        MsgBox "Pilih terlebih dahulu agent awalnya!", vbOKOnly + vbInformation, "Informasi"
        CmbAgentCollBersama.SetFocus
        Exit Sub
    End If

    M_WHERE = ""
    GroupingTL = ""
    
    cmdsql = "select * from mgm "
    
    If CmbAgentCollBersama.text = "ALL" Then
        
        'Ini jika agent=ALL dan status account=ALL
        If CmbStatusCollBersama.text = "ALL" Then
            'CUKUP DEH NGGA USAH PAKE SCRIPT---------
        
        '@@15022013 Ini jika filter agentnya=ALL tetapi status accountnya <> ALL
        ElseIf CmbStatusCollBersama.text <> "ALL" Then
            If CmbStatusCollBersama.text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusCollBersama.text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusCollBersama.text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusCollBersama.text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusCollBersama.text <> "LPD 1" And _
                   CmbStatusCollBersama.text <> "LPD 2" And _
                   CmbStatusCollBersama.text <> "LPD 3" And _
                   CmbStatusCollBersama.text <> "LPD 3<" Then
                   
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusCollBersama.text & "%' "
                Else
                    M_WHERE = " and f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusCollBersama.text & "%' "
                End If
            End If
        End If
        
    ElseIf CmbAgentCollBersama.text <> "ALL" Then
    
        '@@19022013 Tambahan ini buat grouping TL
        If UCase(Mid(CmbAgentCollBersama.text, 1, 2)) = "TL" Then
            GroupingTL = " agent_asli in (select userid from usertbl where spvcode in ("
            GroupingTL = GroupingTL & " select spvcode from usertbl where userid='"
            GroupingTL = GroupingTL & CmbAgentCollBersama.text + "')) "
        Else
            GroupingTL = " agent_asli='"
            GroupingTL = GroupingTL + CmbAgentCollBersama.text + "' "
        End If
                
        'Ini jika agent <>ALL dan status account=ALL
        If CmbStatusCollBersama.text = "ALL" Then
            If M_WHERE = "" Then
                'M_WHERE = " where agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " where " & GroupingTL
            Else
                'M_WHERE = " and agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " and " & GroupingTL
            End If
            
        'Ini jika agent=<>ALL dan status account <> ALL
        ElseIf CmbStatusCollBersama.text <> "ALL" Then
            If CmbStatusCollBersama.text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusCollBersama.text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and  " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            End If
            
            If CmbStatusCollBersama.text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusCollBersama.text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusCollBersama.text <> "LPD 1" And _
                   CmbStatusCollBersama.text <> "LPD 2" And _
                   CmbStatusCollBersama.text <> "LPD 3" And _
                   CmbStatusCollBersama.text <> "LPD 3<" Then
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusCollBersama.text & "%' and " & GroupingTL
                Else
                    M_WHERE = " and f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusCollBersama.text & "%' and " & GroupingTL
                End If
            End If
        End If
            
    End If
    
    If M_WHERE = "" Then
        M_WHERE = " where agent not in ('LUNAS','COMPLAIN','CLAIM') and agent='AKSESALL' order by name asc "
    Else
        M_WHERE = M_WHERE & " and agent not in ('LUNAS','COMPLAIN','CLAIM') and agent='AKSESALL' order by name asc "
    End If
    
    
    
    DoEvents
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql + M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.clear
    TxtJmlhAcc.text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set listItem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            listItem.SubItems(1) = M_Objrs("name")
            listItem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            listItem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            listItem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            listItem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            listItem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                listItem.ForeColor = vbRed
                listItem.ListSubItems(1).ForeColor = vbRed
                listItem.ListSubItems(2).ForeColor = vbRed
                listItem.ListSubItems(3).ForeColor = vbRed
                listItem.ListSubItems(4).ForeColor = vbRed
                listItem.ListSubItems(5).ForeColor = vbRed
                listItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                listItem.ForeColor = vbBlue
                listItem.ListSubItems(1).ForeColor = vbBlue
                listItem.ListSubItems(2).ForeColor = vbBlue
                listItem.ListSubItems(3).ForeColor = vbBlue
                listItem.ListSubItems(4).ForeColor = vbBlue
                listItem.ListSubItems(5).ForeColor = vbBlue
                listItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub cmd_insert_antrian_Click()
    Dim iQuery As String
    Dim CustId, nama_customer, status_account, agent_saat_ini, agent_terdahulu, kode_profile As String
    Dim agent, tanggal_awal, waktu_awal, tanggal_akhir, waktu_akhir As String
    Dim i As Integer
    Dim S As Long
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data Customer Tidak Tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtAgent.text = "" Then
        MsgBox "Agent Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglAwal.ValueIsNull Then
        MsgBox "Tanggal Mulai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtWaktuAwal.ValueIsNull Then
        MsgBox "Waktu Mulai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglAkhir.ValueIsNull Then
        MsgBox "Tanggal Selesai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtWaktuAkhir.ValueIsNull Then
        MsgBox "Waktu Selesai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    M_OBJCONN.execute "DELETE FROM temp_approval_aksesall"
    
    PB1.Max = LvAcc.ListItems.Count
    For S = 1 To LvAcc.ListItems.Count
        PB1.Value = S
        Me.MousePointer = vbHourglass
        Me.Enabled = False
        CustId = LvAcc.ListItems(S).text
        nama_customer = LvAcc.ListItems(S).ListSubItems(1)
        status_account = LvAcc.ListItems(S).ListSubItems(2)
        agent_saat_ini = LvAcc.ListItems(S).ListSubItems(3)
        agent_terdahulu = LvAcc.ListItems(S).ListSubItems(4)
        kode_profile = lbl_profile.Caption
        agent = Replace(TxtAgent.text, "'", "''")
        tanggal_awal = Format(TxtTglAwal.Value, "dd/mm/yyyy")
        waktu_awal = Format(TxtWaktuAwal.Value, "hh:nn")
        tanggal_akhir = Format(TxtTglAkhir.Value, "dd/mm/yyyy")
        waktu_akhir = Format(TxtWaktuAkhir.Value, "hh:nn")

        iQuery = "INSERT INTO temp_approval_aksesall"
        iQuery = iQuery + " VALUES ('" & CustId & "', '" & nama_customer & "', '" & status_account & "', "
        iQuery = iQuery + " '" & agent_saat_ini & "', '" & agent_terdahulu & "',  '" & kode_profile & "',  "
        iQuery = iQuery + " '" & agent & "', '" & tanggal_awal & "', '" & waktu_awal & "', '" & tanggal_akhir & "', "
        iQuery = iQuery + " '" & waktu_akhir & "')"
        
        M_OBJCONN.execute iQuery
    Next S
     
    MsgBox "Aksesall Berhasil Diproses, Menunggu Approve Dari SPV Atau Manager!", vbOKOnly + vbInformation, "Informasi"
        Me.MousePointer = vbNormal
        Me.Enabled = True
    Unload Me

End Sub

Private Sub CmdBukaAccount_Click()
    Dim cmdsql As String
    Dim w, K, S As Integer
    Dim a As String
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data account tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin membuka lock account yang terceklist?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    For K = 1 To LvAcc.ListItems.Count
        If LvAcc.ListItems(K).Checked = True Then
            S = S + 1
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Anda belum menceklist account yang akan dibuka!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvAcc.ListItems.Count
    For w = 1 To LvAcc.ListItems.Count
        PB1.Value = w
        If LvAcc.ListItems(w).Checked = True Then
            'buka locknya
            cmdsql = "update mgm set monitor_akses=null,waktu_akses=null where custid='"
            cmdsql = cmdsql & CStr(LvAcc.ListItems(w).text) & "'"
            M_OBJCONN.execute cmdsql
        End If
    Next w
    
    Call IsiAccount
    
    MsgBox "Proses berhasil!", vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdCari_Click()
    Call IsiAccount
End Sub

Private Sub CmdCekAllAcc_Click()
    Dim w As Long
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvAcc.ListItems.Count
        LvAcc.ListItems(w).Checked = True
    Next w
    TxtJmlhAcc.text = LvAcc.ListItems.Count
End Sub

Private Sub CmdCekAllAgent_Click()
    Dim w As Integer
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvAgent.ListItems.Count
        LvAgent.ListItems(w).Checked = True
    Next w
End Sub

Private Sub CmdClear_Click()
    TxtCariCustid.text = ""
    TxtCariNama.text = ""
End Sub

Private Sub CmdEdit_Click()
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data Account tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data Agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    FrmEditDistribusiAccount.TxtID.text = LvAgent.SelectedItem.text
    FrmEditDistribusiAccount.TxtAgent.text = LvAgent.SelectedItem.SubItems(1)
    FrmEditDistribusiAccount.Show vbModal
End Sub

Private Sub CmdFilterExcel_Click()
    FrmFilterExcelDistribusiAcc.Show vbModal
End Sub

Private Sub CmdFormClaimAccount_Click()
    FrmListClaim.Show vbModal
End Sub

Private Sub CmdHapusAgent_Click()
    Dim a As String
    Dim cmdsql As String
    Dim w, i, K As Integer
    Dim M_Objrs As ADODB.Recordset
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data agent akan dihapus?", vbYesNo + vbInformation, "Informasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    i = 0
    For w = 1 To LvAgent.ListItems.Count
       If LvAgent.ListItems(w).Checked = True Then
            i = i + 1
       End If
    Next w
    
    
    If i = 0 Then
        MsgBox "Anda belum memilih data agent yang akan dihapus!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
        
    DoEvents
        
    PB1.Max = LvAgent.ListItems.Count
        
    For K = 1 To LvAgent.ListItems.Count
        PB1.Value = K
        If LvAgent.ListItems(K).Checked = True Then
            cmdsql = "DELETE FROM tbl_cust_not_aksesall WHERE kd_profile='" & LvAgent.ListItems(K).SubItems(7) & "' " & _
                    "AND custid='" & LvAgent.ListItems(K).SubItems(2) & "' AND agent='" & LvAgent.ListItems(K).SubItems(1) & "'"
            M_OBJCONN.execute cmdsql
            
            cmdsql = "INSERT INTO tbl_cust_not_aksesall(kd_profile,custid,agent) " & _
                    "VALUES('" & LvAgent.ListItems(K).SubItems(7) & "','" & LvAgent.ListItems(K).SubItems(2) & "' & LvAgent.ListItems(K).SubItems(1) & " ')"
'            Cmdsql = "delete from tbl_distribusi_account where id='"
'            Cmdsql = Cmdsql + CStr(LvAgent.ListItems(K).Text) + "'"
            M_OBJCONN.execute cmdsql
'
'            'Update status agentnya nih
'            Cmdsql = "update usertbl set f_akses_all_acc=null,f_pesanresetauto='1' "
'            Cmdsql = Cmdsql + " where userid='"
'            Cmdsql = Cmdsql + CStr(LvAgent.ListItems(K).SubItems(1)) + "'"
'            M_OBJCONN.Execute Cmdsql
'
'            'Cek apakah custid ini sudah habis agentnya?
'            Cmdsql = "select * from tbl_distribusi_account where custid='"
'            Cmdsql = Cmdsql & CStr(LvAgent.ListItems(K).SubItems(2)) & "'"
'            Set M_Objrs = New ADODB.Recordset
'            M_Objrs.CursorLocation = adUseClient
'            M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If M_Objrs.RecordCount = 0 Then
'                'Update ke agent yang lama
'                Cmdsql = "update mgm set agent=agent_asli,agent_asli=null,"
'                Cmdsql = Cmdsql + " user_claim=null,waktu_claim=null,alasan_claim=null "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql & CStr(LvAgent.ListItems(K).SubItems(2)) + "' and agent_asli is not null "
'                M_OBJCONN.Execute Cmdsql
'            End If
'            Set M_Objrs = Nothing
            
        End If
    Next K
    
    'Cek apakah custid ini sudah habis agentnya?
'    Cmdsql = "select * from tbl_distribusi_account where custid='"
'    Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
'    If M_Objrs.RecordCount = 0 Then
'        If LvAcc.SelectedItem.SubItems(3) <> "" Then
'            A = MsgBox("Account ini sudah tidak ada yang memiliki. Anda ingin mengembalikannya ke agent terdahulu?", vbYesNo + vbQuestion, "Konfirmasi")
'            If A = vbYes Then
'                'Update ke agent yang lama
'                Cmdsql = "update mgm set agent=agent_asli,agent_asli=null where custid='"
'                Cmdsql = Cmdsql & CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'            Else
'                'Update ke agent yang kosong
'                Cmdsql = "update mgm set agent='#KOSONG#' where custid='"
'                Cmdsql = Cmdsql & CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'            End If
'        Else
'            'Update ke agent yang kosong
'            Cmdsql = "update mgm set agent='#KOSONG#' where custid='"
'            Cmdsql = Cmdsql & CStr(lblCustId.Caption) + "'"
'            M_OBJCONN.Execute Cmdsql
'        End If
'    End If
    
    Call CariAgent
    
    MsgBox "Data agent berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
    
    'Call IsiAccount
End Sub

Private Sub CmdKembalikanAgent_Click()
    Dim cmdsql As String
    Dim w, K, S As Integer
    Dim a As String
    Dim M_Objrs As ADODB.Recordset
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data account tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan mengembalikan account yang diceklist ke agent awal?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    For K = 1 To LvAcc.ListItems.Count
        If LvAcc.ListItems(K).Checked = True Then
            S = S + 1
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Anda belum menceklist account yang akan dikembalikan agentnya!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    
    PB1.Max = LvAcc.ListItems.Count
    For w = 1 To LvAcc.ListItems.Count
        PB1.Value = w
        If LvAcc.ListItems(w).Checked = True Then
            If M_Objrs.state = 1 Then M_Objrs.Close
            cmdsql = "SELECT * FROM tbllog_claim_aksesall WHERE custid='" & CStr(LvAcc.ListItems(w).text) & "'"
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs.RecordCount = 0 Then
                ' 19 AGUSTUS 2014 agent_asli=null dihilangkan
                cmdsql = "update mgm set agent=agent_asli where custid='"
                cmdsql = cmdsql + CStr(LvAcc.ListItems(w).text) + "' and agent_asli is not null "
                M_OBJCONN.execute cmdsql
            End If
            'Hapus data di tabel distribusinya
'            Cmdsql = "delete from tbl_distribusi_account where custid='"
'            Cmdsql = Cmdsql + CStr(LvAcc.ListItems(W).Text) + "'"
            cmdsql = "DELETE FROM tbl_cust_aksesall WHERE custid='" & CStr(LvAcc.ListItems(w).text) & "'"
            M_OBJCONN.execute cmdsql
        End If
    Next w
    
    Set M_Objrs = Nothing
    
    Call IsiAccount
    
    MsgBox "Account berhasil dikembalikan ke agent awal!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdLihatListAgent_Click()
    FrmListAgent.Show vbModal
End Sub

Private Sub CmdProses_Click()
    Dim cmdsql, AmbilCustid As String
    Dim M_Objrs As ADODB.Recordset
    Dim w, K, S As Long
    Dim Tanggal1 As String
    Dim Tanggal2 As String
    Dim pesan As String
    Dim a As String
    Dim M_ObjrsWaktuServer As ADODB.Recordset
    Dim WaktuServer As String
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data customer tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtAgent.text = "" Then
        MsgBox "Agent tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    For w = 1 To LvAcc.ListItems.Count
        If LvAcc.ListItems(w).Checked = True Then
            S = S + 1
        End If
    Next w
    
    If S = 0 Then
        MsgBox "Anda belum memilih data customer!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin menandai account dapat di collect bersama?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
        
    MDIForm1.Timer1 = False
    MDIForm1.TimerCTI = False
    MDIForm1.TimerBlink = False
    'Cek waktu server
    cmdsql = "select now()"
    Set M_ObjrsWaktuServer = New ADODB.Recordset
    M_ObjrsWaktuServer.CursorLocation = adUseClient
    DoEvents
    M_ObjrsWaktuServer.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    'WaktuServer = Format(M_ObjrsWaktuServer(0), "m/dd/yyyy hh:nn:ss")
    WaktuServer = Format(M_ObjrsWaktuServer(0), "yyyy-mm-dd hh:nn:ss")
    
    

    If TxtTglAwal.ValueIsNull = True Or _
       TxtWaktuAwal.ValueIsNull = True Or _
       TxtTglAkhir.ValueIsNull = True Or _
       TxtWaktuAkhir.ValueIsNull = True Or _
       TxtTglExpired.ValueIsNull = True Then
        MsgBox "Waktu tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Cek tanggal awal tidak boleh lebih besar dari tanggal akhir
'    Tanggal1 = Format(TxtTglAwal.Value, "m/dd/yyyy") & " " & Format(TxtWaktuAwal.Value, "hh:nn")
'    Tanggal2 = Format(TxtTglAkhir.Value, "m/dd/yyyy") & " " & Format(TxtWaktuAkhir.Value, "hh:nn")
    Tanggal1 = Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value, "hh:nn")
    Tanggal2 = Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value, "hh:nn")
     
    'Cek jika waktu akhir server lebih kecil dari waktu server sekarang
    If CDate(Tanggal2) < CDate(WaktuServer) Then
        MsgBox "Waktu akhir tidak boleh lebih kecil dari waktu server! Waktu Server sekarng: " & Format(WaktuServer, "yyyy-mm-dd hh:nn:ss"), vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
     
    If CDate(Tanggal1) > CDate(Tanggal2) Then
        MsgBox "Tanggal awal tidak boleh lebih besar dari tanggal akhir!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'On Error GoTo SALAH
    
    'Ambil Data agent
    cmdsql = "select * from usertbl where userid in ("
    cmdsql = cmdsql + TxtAgent.text + ")"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If M_Objrs.RecordCount > 0 Then
        DoEvents
        LvAcc.Enabled = False
        CmdCari.Enabled = False
        cmdClear.Enabled = False
        CmdLihatListAgent.Enabled = False
        TxtTglAwal.Enabled = False
        TxtWaktuAwal.Enabled = False
        TxtTglAkhir.Enabled = False
        TxtWaktuAkhir.Enabled = False

        CmbAgentCollBersama.Enabled = False
        CmbStatusCollBersama.Enabled = False

        While Not M_Objrs.EOF
            DoEvents
            'Update status Agentnya
            cmdsql = "update usertbl set f_akses_all_acc='1',f_pesanresetauto='1',profile_akses_all='" & lbl_profile.Caption & "' "
            cmdsql = cmdsql + " where userid='"
            cmdsql = cmdsql + M_Objrs("userid") + "'"
            M_OBJCONN.execute cmdsql

            'Kirim pesan ke agent
            pesan = "Pesan dibuat otomatis oleh system!" & vbCrLf
            pesan = pesan & "----------------------------------------------" & vbCrLf
            pesan = pesan & "SPV menambahkan account baru untuk anda. " & vbCrLf
            pesan = pesan & "Account ini dapat di collect secara bersama-sama oleh anda, " & vbCrLf
            pesan = pesan & "mulai dari :" & Format(Tanggal1, "yyyy-mm-dd hh:nn:ss") & " s.d. " & vbCrLf
            pesan = pesan & Format(Tanggal2, "yyyy-mm-dd hh:nn:ss") & vbCrLf
            pesan = pesan & "Cek account baru anda dengan mengklik ulang tombol search data!"

            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + M_Objrs("userid") + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.Text1.text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + pesan + "')"

            M_OBJCONN.execute cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
    
    ' Hapus Header
    cmdsql = "delete from tbl_profile_aksesall where kd_profile='" & lbl_profile.Caption & "'"
    M_OBJCONN.execute cmdsql
    
    ' Hapus detail
    cmdsql = "delete from tbl_cust_aksesall where kd_profile='" & lbl_profile.Caption & "'"
    M_OBJCONN.execute cmdsql

    ' Insert Header
    cmdsql = "INSERT INTO tbl_profile_aksesall(kd_profile,waktu_awal,waktu_akhir,log_distribusi,log_tgl_distribusi) VALUES"
    cmdsql = cmdsql & "('" & lbl_profile.Caption & "', '"
    cmdsql = cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
    cmdsql = cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
    cmdsql = cmdsql & MDIForm1.Text1.text & "',"
    cmdsql = cmdsql & " now())"
    M_OBJCONN.execute cmdsql
    
    PB1.Max = LvAcc.ListItems.Count
    For K = 1 To LvAcc.ListItems.Count
        PB1.Value = K
        DoEvents
        If LvAcc.ListItems(K).Checked = True Then
            ' DELETE DATA CUSTOMER DI PROFILE SEBELUMNYA KLO ADA
            cmdsql = "DELETE FROM tbl_cust_aksesall WHERE custid='" & CStr(LvAcc.ListItems(K).text) & "'"
            M_OBJCONN.execute cmdsql
            
            ' UPDATE MGM TGL EXPIRED CLAIM
            cmdsql = "update mgm set tgl_exp_claim = '" & Format(TxtTglExpired.Value, "yyyy-mm-dd") & "' where custid = '" & CStr(LvAcc.ListItems(K).text) & "'"
            M_OBJCONN.execute cmdsql
            
            'Inputkan data ke Detail
            cmdsql = "INSERT INTO tbl_cust_aksesall values('" & lbl_profile.Caption & "','" & CStr(LvAcc.ListItems(K).text) & "')"
            M_OBJCONN.execute cmdsql
            
            If UCase(LvAcc.ListItems(K).SubItems(3)) = "#KOSONG#" Then
                cmdsql = "UPDATE mgm SET agent='AKSESALL' WHERE custid='"
                cmdsql = cmdsql & CStr(LvAcc.ListItems(K).text) & "'"
                M_OBJCONN.execute cmdsql
            ElseIf UCase(LvAcc.ListItems(K).SubItems(3)) <> "AKSESALL" Then
                cmdsql = "UPDATE mgm SET agent_asli=agent, agent='AKSESALL' WHERE custid='"
                cmdsql = cmdsql & CStr(LvAcc.ListItems(K).text) & "'"
                M_OBJCONN.execute cmdsql
            End If
            
            ' ====== UPDATE IZUDDIN 08 OKTOBER 2013 =======
            cmdsql = "SELECT custid,agent_asli, agent FROM mgm WHERE custid='" & CStr(LvAcc.ListItems(K).text) & "'"
            
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs.RecordCount > 0 Then
                M_OBJCONN.execute "INSERT INTO tbl_hst_aksesall(custid,agent_asli) values('" & M_Objrs!CustId & "','" & M_Objrs!agent_asli & "')"
            End If
            
            Set M_Objrs = Nothing
            ' =============================================
        End If
        Debug.Print "loop ke" & CStr(K)
    Next K
    
    'Ambil Data agent
'    Cmdsql = "select * from usertbl where userid in ("
'    Cmdsql = Cmdsql + TxtAgent.Text + ")"
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_Objrs.RecordCount > 0 Then
'        DoEvents
'        LvAcc.Enabled = False
'        CmdCari.Enabled = False
'        cmdClear.Enabled = False
'        CmdLihatListAgent.Enabled = False
'        TxtTglAwal.Enabled = False
'        TxtWaktuAwal.Enabled = False
'        TxtTglAkhir.Enabled = False
'        TxtWaktuAkhir.Enabled = False
'
'        CmbAgentCollBersama.Enabled = False
'        CmbStatusCollBersama.Enabled = False
'
'        PB1.Max = M_Objrs.RecordCount
'        While Not M_Objrs.EOF
'            PB1.Value = M_Objrs.Bookmark
'            DoEvents
'            'Update status Agentnya
'            Cmdsql = "update usertbl set f_akses_all_acc='1',f_pesanresetauto='1' "
'            Cmdsql = Cmdsql + " where userid='"
'            Cmdsql = Cmdsql + M_Objrs("userid") + "'"
'            M_OBJCONN.Execute Cmdsql
'
'            'Kirim pesan ke agent
'            Pesan = "Pesan dibuat otomatis oleh system!" & vbCrLf
'            Pesan = Pesan & "----------------------------------------------" & vbCrLf
'            Pesan = Pesan & "SPV menambahkan account baru untuk anda. " & vbCrLf
'            Pesan = Pesan & "Account ini dapat di collect secara bersama-sama oleh anda, " & vbCrLf
'            Pesan = Pesan & "mulai dari :" & Format(Tanggal1, "yyyy-mm-dd hh:nn:ss") & " s.d. " & vbCrLf
'            Pesan = Pesan & Format(Tanggal2, "yyyy-mm-dd hh:nn:ss") & vbCrLf
'            Pesan = Pesan & "Cek account baru anda dengan mengklik ulang tombol search data!"
'
'            Cmdsql = "insert into msgtbl "
'            Cmdsql = Cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
'            Cmdsql = Cmdsql + M_Objrs("userid") + "','"
'            Cmdsql = Cmdsql + Format(Now(), "yyyymmdd") + "','"
'            Cmdsql = Cmdsql + MDIForm1.Text1.Text + "','"
'            Cmdsql = Cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
'            Cmdsql = Cmdsql + Pesan + "')"
'
'            M_OBJCONN.Execute Cmdsql
'
'
'            For K = 1 To LvAcc.ListItems.Count
'                DoEvents
'                If LvAcc.ListItems(K).Checked = True Then
'                    'Hapus dulu jika ada data sebelumnya
'                    Cmdsql = "delete from tbl_distribusi_account where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "' and agent='"
'                    Cmdsql = Cmdsql & M_Objrs("userid") & "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                    'Inputkan data ke database
'                    Cmdsql = "insert into tbl_distribusi_account (custid,agent,"
'                    Cmdsql = Cmdsql & "waktu_awal,waktu_akhir, log_distribusi,"
'                    Cmdsql = Cmdsql & "log_tgl_distribusi) values ('"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) + "','"
'                    Cmdsql = Cmdsql & M_Objrs("userid") & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
'                    Cmdsql = Cmdsql & MDIForm1.Text1.Text & "',"
'                    Cmdsql = Cmdsql & " now())"
'                    M_OBJCONN.Execute Cmdsql
'
'                End If
'                Debug.Print "loop ke" & CStr(K)
'            Next K
'            Debug.Print "record ke" & M_Objrs.Bookmark
'            M_Objrs.MoveNext
'        Wend
'
'        'Catet agent lama
'        PB1.Max = LvAcc.ListItems.Count
'        For K = 1 To LvAcc.ListItems.Count
'        DoEvents
'            PB1.Value = K
'            If LvAcc.ListItems(K).Checked = True Then
'                '@@12022013 Jika status account sebelumnya #KOSONG# atau AKSESALL, ga usah diupdate
'                If UCase(LvAcc.ListItems(K).SubItems(3)) = "#KOSONG#" Then
'                    Cmdsql = "update mgm set  agent='AKSESALL' where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
'                    M_OBJCONN.Execute Cmdsql
'                ElseIf UCase(LvAcc.ListItems(K).SubItems(3)) <> "AKSESALL" Then
'                    Cmdsql = "update mgm set agent_asli=agent, agent='AKSESALL' where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                    '@@18022013 Ini buat, inputin otomatis buat pemilik agent masuk juga akses all
'                    'Hapus dulu jika ada data sebelumnya
'                    Cmdsql = "delete from tbl_distribusi_account where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "' and agent='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).SubItems(3)) & "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                    'Inputkan data ke database
'                    Cmdsql = "insert into tbl_distribusi_account (custid,agent,"
'                    Cmdsql = Cmdsql & "waktu_awal,waktu_akhir, log_distribusi,"
'                    Cmdsql = Cmdsql & "log_tgl_distribusi) values ('"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) + "','"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).SubItems(3)) & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
'                    Cmdsql = Cmdsql & MDIForm1.Text1.Text & "',"
'                    Cmdsql = Cmdsql & " now())"
'                    M_OBJCONN.Execute Cmdsql
'
'                    'Update statusnya
'                    Cmdsql = "update usertbl set f_akses_all_acc='1', f_pesanresetauto='1' where "
'                    Cmdsql = Cmdsql + " userid='"
'                    Cmdsql = Cmdsql + CStr(LvAcc.ListItems(K).SubItems(3)) + "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                End If
'            End If
'            Debug.Print "SET agent ke " & CStr(K)
'        Next K
'    End If
    
    LvAcc.Enabled = True
    CmdCari.Enabled = True
    cmdClear.Enabled = True
    CmdLihatListAgent.Enabled = True
    TxtTglAwal.Enabled = True
    TxtWaktuAwal.Enabled = True
    TxtTglAkhir.Enabled = True
    TxtWaktuAkhir.Enabled = True
    
    CmbAgentCollBersama.Enabled = True
    CmbStatusCollBersama.Enabled = True
    
    'Call IsiAccount
    CmbStatusAcc_Click
    IsiAgentCollectBersama
    IsiStatusCollectBersama
    
    'sessionaksesall
    qc = "create table tblsessaksesall_" & lbl_profile.Caption & " ( position smallint );" & vbCrLf
    qi = qc + "INSERT INTO tblsessaksesall_" & lbl_profile.Caption & " values (1);"
    Mix = qi
    
    M_OBJCONN.execute Mix
    
    
    Call new_kdprofile
    
    ' Hidupkan TIMER
    MDIForm1.Timer1 = True
    MDIForm1.TimerCTI = True
    MDIForm1.TimerBlink = True
    
    MsgBox "Proses berhasil!", vbOKOnly + vbInformation, "Informasi"
    M_OBJCONN.execute "DELETE FROM temp_approval_aksesall"
    Label25.Visible = False
    
    Exit Sub
'SALAH:
'    MsgBox "Mohon maaf ada kesalahan: " & err.Description, "Error"
End Sub

Private Sub CmdUnCekAll_Click()
    Dim w As Integer
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvAcc.ListItems.Count
        LvAcc.ListItems(w).Checked = False
    Next w
    TxtJmlhAcc.text = 0
End Sub

Private Sub CmdUncekallAgent_Click()
        Dim w As Integer
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvAgent.ListItems.Count
        LvAgent.ListItems(w).Checked = False
    Next w
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim sqlfilter As String
    Dim pil_check As Boolean
    Dim pil_batch As Boolean
    Dim list_sts As String
    Dim list_btch As String
    Dim sts_arr() As String
    Dim sts_WO As String
    
    pil_check = False
    
    sqlfilter = ""
    
'    If cb_batch.Text <> "" Then
'        sqlfilter = " AND trim(RECSOURCE)='" & Trim(cb_batch.Text) & "' "
'    End If

    For i = 0 To list_batch.ListCount - 1
        If list_batch.Selected(i) = True Then
            pil_batch = True
            list_btch = list_btch & "'" & list_batch.list(i) & "',"
        End If
    Next i
    
    If pil_batch = True Then
        list_btch = Mid(list_btch, 1, Len(list_btch) - 1)
        sqlfilter = " AND trim(RECSOURCE) in (" & list_btch & ") "
    End If
    
    If IsDate(tgl_lpd.Value) And IsDate(tgl_lpd2.Value) Then
        If tgl_lpd.Value < tgl_lpd2.Value Then
            sqlfilter = sqlfilter & " AND ( CASE WHEN pay_dt_update IS NOT NULL THEN " & _
                        "date(pay_dt_update) between '" & Format(tgl_lpd.Value, "yyyy-mm-01") & "' AND '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(tgl_lpd2.Value, "yyyy-mm-01"))), "yyyy-mm-dd") & "' " & _
                        "ELSE date(pay_dt) between '" & Format(tgl_lpd.Value, "yyyy-mm-01") & "' AND '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(tgl_lpd2.Value, "yyyy-mm-01"))), "yyyy-mm-dd") & "' END ) "
        Else
            MsgBox "LPD 1 harus lebih kecil dari LPD 2", vbCritical + vbInformation, "INFO"
            Exit Sub
        End If
    End If
    
    list_sts = ""
    For i = 0 To Check1.UBound
        If Check1(i).Value = 1 Then
            pil_check = True
            sts_arr = Split(Check1(i).Caption, "-")
            list_sts = list_sts & "'" & sts_arr(0) & "-',"
        End If
    Next i
    
    If pil_check = True Then
        list_sts = Mid(list_sts, 1, Len(list_sts) - 1)
        sqlfilter = sqlfilter & " AND substring(f_cek_new,1,3) in (" & list_sts & ") "
    End If
    
    sts_WO = ""
    If Check2.Value = 1 Then
        For i = List1.ListCount - 1 To 0 Step -1
            If List1.Selected(i) = True Then
                sts_WO = sts_WO & List1.list(i) & ","
            End If
        Next i
        
        If Trim(sts_WO) <> "" Then
            sts_WO = Mid(sts_WO, 1, Len(sts_WO) - 1)
            sqlfilter = sqlfilter & " AND (date_part('year',b_d) in (" & sts_WO & ")) "
        End If
    End If
    
    If Check3.Value = 1 Then
        For i = 0 To SSOption1.UBound
            If SSOption1(i).Value = True Then
                sts_arr = Split(SSOption1(i).Caption, "-")
                sqlfilter = sqlfilter & " AND (curbal >=" & Replace(Trim(sts_arr(0)), ".", "") & " AND curbal <= " & Replace(Trim(sts_arr(1)), ".", "") & " ) "
            End If
        Next i
    End If
    
    ' 20 AGUSTUS 2014 - Review tidak di akses all
    cmdsql = "SELECT * FROM mgm WHERE custid is not null " & sqlfilter
    cmdsql = cmdsql & " AND agent NOT IN ('LUNAS','COMPLAIN','CLAIM','AKSESALL','REVIEW','REVIEW1','REVIEW2','REVIEW3','REVIEW4','REVIEW5','REVIEW6','REVIEW7','REVIEW8','REVIEW9','REVIEW10') AND coalesce(agent,'')<>'' "
    cmdsql = cmdsql & " AND custid NOT IN (select distinct custid from tblsendptp ) "
    ' TAMBAHAN AGAR CLASS 835 TIDAK KENA AKSES ALL
    If Check_decease.Value = 0 Then
        cmdsql = cmdsql & " AND coalesce(cust_class,'')<>'835' "
    End If
    ' -------------------------------------------
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LvAcc.ListItems.clear
    If M_Objrs.RecordCount > 0 Then
        With Me
            .PB1.Max = M_Objrs.RecordCount
            While Not M_Objrs.EOF
                .PB1.Value = M_Objrs.Bookmark
                Set listItem = .LvAcc.ListItems.ADD(, , M_Objrs("custid"))
                listItem.SubItems(1) = M_Objrs("name")
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                listItem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
                If UCase(M_Objrs("agent")) = "AKSESALL" Then
                    listItem.ForeColor = vbRed
                    listItem.ListSubItems(1).ForeColor = vbRed
                    listItem.ListSubItems(2).ForeColor = vbRed
                    listItem.ListSubItems(3).ForeColor = vbRed
                    listItem.ListSubItems(4).ForeColor = vbRed
                    listItem.ListSubItems(5).ForeColor = vbRed
                    listItem.ListSubItems(6).ForeColor = vbRed
                End If
            
                If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                    listItem.ForeColor = vbBlue
                    listItem.ListSubItems(1).ForeColor = vbBlue
                    listItem.ListSubItems(2).ForeColor = vbBlue
                    listItem.ListSubItems(3).ForeColor = vbBlue
                    listItem.ListSubItems(4).ForeColor = vbBlue
                    listItem.ListSubItems(5).ForeColor = vbBlue
                    listItem.ListSubItems(6).ForeColor = vbBlue
                End If
                M_Objrs.MoveNext
            Wend
        End With
        MsgBox "Data berhasil di load!", vbOKOnly + vbInformation, "Informasi"
    Else
        MsgBox "Data tidak ditemukan !", vbOKOnly + vbInformation, "Info"
    End If
    Set M_Objrs = Nothing
End Sub

Private Sub Command2_Click()
    Frame1.Visible = True
End Sub

Private Sub Command3_Click()
    Frame1.Visible = False
End Sub

Private Sub NgisiDataAksesallPending()
    Dim sQuery As String
    Dim Randy_RS As ADODB.Recordset
    
    sQuery = "SELECT * FROM temp_approval_aksesall"
    Set Randy_RS = New ADODB.Recordset
    Randy_RS.CursorLocation = adUseClient
    Randy_RS.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Randy_RS.RecordCount > 0 Then
        TxtAgent.text = IIf(IsNull(Randy_RS("agent")), "", Randy_RS("agent"))
        TxtTglAwal.text = Format(IIf(IsNull(Randy_RS("tanggal_awal")), "", Randy_RS("tanggal_awal")), "YYYY-MM-DD")
        TxtWaktuAwal.text = IIf(IsNull(Randy_RS("waktu_awal")), "", Randy_RS("waktu_awal"))
        TxtTglAkhir.text = Format(IIf(IsNull(Randy_RS("tanggal_akhir")), "", Randy_RS("tanggal_akhir")), "YYYY-MM-DD")
        TxtWaktuAkhir.text = IIf(IsNull(Randy_RS("waktu_akhir")), "", Randy_RS("waktu_akhir"))
    End If
    
     With FrmDistribusiAcc
        .LvAcc.ListItems.clear
        .PB1.Max = Randy_RS.RecordCount
        While Not Randy_RS.EOF
            .PB1.Value = Randy_RS.Bookmark
            Set listItem = .LvAcc.ListItems.ADD(, , Randy_RS("custid"))
            listItem.SubItems(1) = Randy_RS("nama_customer")
            listItem.SubItems(2) = IIf(IsNull(Randy_RS("status_account")), "", Randy_RS("status_account"))
            listItem.SubItems(3) = IIf(IsNull(Randy_RS("agent_saat_ini")), "", Randy_RS("agent_saat_ini"))
            listItem.SubItems(4) = IIf(IsNull(Randy_RS("agent_terdahulu")), "", Randy_RS("agent_terdahulu"))
                        
            If UCase(Randy_RS("agent")) = "AKSESALL" Then
                listItem.ForeColor = vbRed
                listItem.ListSubItems(1).ForeColor = vbRed
                listItem.ListSubItems(2).ForeColor = vbRed
                listItem.ListSubItems(3).ForeColor = vbRed
                listItem.ListSubItems(4).ForeColor = vbRed
                listItem.ListSubItems(5).ForeColor = vbRed
                listItem.ListSubItems(6).ForeColor = vbRed
            End If
        
            If UCase(Randy_RS("agent")) = "#KOSONG#" Then
                listItem.ForeColor = vbBlue
                listItem.ListSubItems(1).ForeColor = vbBlue
                listItem.ListSubItems(2).ForeColor = vbBlue
                listItem.ListSubItems(3).ForeColor = vbBlue
                listItem.ListSubItems(4).ForeColor = vbBlue
                listItem.ListSubItems(5).ForeColor = vbBlue
                listItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
            Randy_RS.MoveNext
        Wend
    End With
End Sub

Private Sub Command4_Click()
    frmaksesallmonitoring.Show
End Sub

Private Sub Form_Load()
    Dim tglprofile As String
    Dim cmdsql, ran As String
    Dim M_Objrs As ADODB.Recordset
    
    If MDIForm1.Text2.text = "Manager" Then
        CmdProses.Visible = False
    End If
    
    
    Call HeaderAgent
    Call HeaderAccount
    Call IsiComboFilter
    Call IsiComboAgent
    Call IsiComboStatusAcc
    
    If MDIForm1.Text2.text = "Supervisor" Then
        Command4.Visible = True
    End If
    
    cmdsql = "select now()"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        LblWaktuServer.Caption = Format(M_Objrs(0), "yyyy-mm-dd hh:nn")
        tglprofile = Format(M_Objrs(0), "yyyymmdd")
    End If
    
    Call new_kdprofile
    Call isi_batch
    
    Set M_Objrs = Nothing
    
    Call IsiAgentCollectBersama
    Call IsiStatusCollectBersama
    
    If CekAksesallPending = False Then
        CmdProses.Enabled = False
        Label25.Visible = False
    Else
        Label25.Visible = True
        
        ran = MsgBox("Ada Pendingan Aksesall, Apakah Mau Di-Approve?", vbYesNo + vbQuestion, "Konfirmasi")
        If ran = vbNo Then
            CmdProses.Enabled = False
            cmd_insert_antrian.Enabled = True
        Else
            CmdProses.Enabled = True
            cmd_insert_antrian.Enabled = False
            Call NgisiDataAksesallPending
        End If
    End If
    
    
    
End Sub

Private Function CekAksesallPending() As Boolean
    Dim sQuery As String
    Dim Randy_RS As ADODB.Recordset
    
    sQuery = "SELECT custid FROM temp_approval_aksesall limit 1"
    Set Randy_RS = New ADODB.Recordset
    Randy_RS.CursorLocation = adUseClient
    Randy_RS.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Randy_RS.RecordCount > 0 Then
        CekAksesallPending = True
    Else
        CekAksesallPending = False
    End If
    
    Set Randy_RS = Nothing
End Function

Private Sub isi_batch()
    Dim M_Objrs As ADODB.Recordset
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open "SELECT distinct RECSOURCE FROM mgm WHERE RECSOURCE is not null ORDER BY RECSOURCE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    list_batch.clear
    If M_Objrs.RecordCount > 0 Then
        'cb_batch.CLEAR
        'cb_batch.AddItem ""
        Do Until M_Objrs.EOF
            'cb_batch.AddItem cnull(M_Objrs!RECSOURCE)
            list_batch.AddItem cnull(M_Objrs!RECSOURCE)
            M_Objrs.MoveNext
        Loop
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub LvAcc_Click()
    Call CariAgent
End Sub

Public Sub CariAgent()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim M_Objrs2 As ADODB.Recordset
    Dim listItem As listItem
    Dim z As Integer
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Peringatan"
        Exit Sub
    End If
    
    
    lblCustId.Caption = LvAcc.SelectedItem.text
    lblNama.Caption = LvAcc.SelectedItem.SubItems(1)
    LblStatusAcc.Caption = IIf(IsNull(LvAcc.SelectedItem.SubItems(2)), "-", LvAcc.SelectedItem.SubItems(2))
    
'    Cmdsql = "select * from tbl_distribusi_account where custid='"
'    Cmdsql = Cmdsql & CStr(lblcustid.Caption) & "' order by agent asc "
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ' UPDATE 22 MEI 2013 BY IZUDDIN
    cmdsql = "select x.userid as agent,a.kd_profile,b.custid,a.waktu_awal,a.waktu_akhir,a.log_distribusi,a.log_tgl_distribusi from usertbl x,tbl_profile_aksesall a, tbl_cust_aksesall b WHERE a.kd_profile=b.kd_profile AND a.kd_profile=x.profile_akses_all AND b.custid='"
    cmdsql = cmdsql & CStr(lblCustId.Caption) & "' ORDER BY userid"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlhAgent.text = M_Objrs.RecordCount
    LvAgent.ListItems.clear
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
'    Set M_Objrs2 = New ADODB.Recordset
'    M_Objrs2.CursorLocation = adUseClient

    z = 0
    While Not M_Objrs.EOF
'        If M_Objrs2.state = 1 Then M_Objrs2.Close
'        Cmdsql = "SELECT agent FROM tbl_cust_not_aksesall WHERE custid='" & M_Objrs("custid") & "' AND agent='" & M_Objrs("agent") & "' AND kd_profile='" & M_Objrs("kd_profile") & "'"
'        M_Objrs2.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        z = z + 1
        Set listItem = LvAgent.ListItems.ADD(, , z)
            listItem.SubItems(1) = M_Objrs("agent")
            listItem.SubItems(2) = M_Objrs("custid")
            listItem.SubItems(3) = Format(M_Objrs("waktu_awal"), "yyyy-mm-dd hh:nn:ss")
            listItem.SubItems(4) = Format(M_Objrs("waktu_akhir"), "yyyy-mm-dd hh:nn:ss")
            listItem.SubItems(5) = M_Objrs("log_distribusi")
            listItem.SubItems(6) = Format(M_Objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
            listItem.SubItems(7) = M_Objrs("kd_profile")
       M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub IsiComboFilter()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbFilterAcc.clear
    
    CmbFilterAcc.AddItem "ALL"
    
    cmdsql = "select * from usertbl where userid not in ('LUNAS','COMPLAIN','CLAIM') order by userid asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            CmbFilterAcc.AddItem M_Objrs("userid")
            M_Objrs.MoveNext
        Wend
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub CariFilter()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    Dim M_WHERE As String
    
    M_WHERE = ""
    
    cmdsql = "select * from mgm  "
    
    If CmbFilterAcc.text <> "ALL" Then
        If M_WHERE = "" Then
            M_WHERE = " where agent='" + CStr(CmbFilterAcc.text) + "' "
            M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS') "
        Else
            M_WHERE = M_WHERE & " and agent='" + CStr(CmbFilterAcc.text) + "' "
            M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS') "
        End If
    End If
    
    
    If M_WHERE <> "" Then
        M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS') "
    Else
        M_WHERE = " where agent not in ('COMPLAIN','LUNAS') "
    End If
    
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql & M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.clear
    TxtJmlhAcc.text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set listItem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            listItem.SubItems(1) = M_Objrs("name")
            listItem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            listItem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            listItem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            listItem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            listItem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                listItem.ForeColor = vbRed
                listItem.ListSubItems(1).ForeColor = vbRed
                listItem.ListSubItems(2).ForeColor = vbRed
                listItem.ListSubItems(3).ForeColor = vbRed
                listItem.ListSubItems(4).ForeColor = vbRed
                listItem.ListSubItems(5).ForeColor = vbRed
                listItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                listItem.ForeColor = vbBlue
                listItem.ListSubItems(1).ForeColor = vbBlue
                listItem.ListSubItems(2).ForeColor = vbBlue
                listItem.ListSubItems(3).ForeColor = vbBlue
                listItem.ListSubItems(4).ForeColor = vbBlue
                listItem.ListSubItems(5).ForeColor = vbBlue
                listItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub LvAcc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAcc.SortKey = ColumnHeader.Index - 1
    LvAcc.Sorted = True
End Sub

Private Sub LvAgent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAgent.SortKey = ColumnHeader.Index - 1
    LvAgent.Sorted = True
End Sub

Private Sub LvAgent_DblClick()
    CmdEdit_Click
End Sub

Private Sub IsiComboAgent()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbAgent.clear
    CmbAgent.AddItem "ALL"
    
    cmdsql = "select * from usertbl where usertype in ('1','6') and userid "
    cmdsql = cmdsql & " not in ('LUNAS','COMPLAIN','COMPLAIN','CLAIM') and userid is not null order by userid asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            CmbAgent.AddItem M_Objrs("userid")
            M_Objrs.MoveNext
        Wend
    End If
    
    Set M_Objrs = Nothing
End Sub

'@@14022013 Tambahan filter status account
Private Sub IsiComboStatusAcc()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbStatusAcc.clear
    CmbStatusAcc.AddItem "ALL"
    CmbStatusAcc.AddItem "LPD 1"
    CmbStatusAcc.AddItem "LPD 2"
    CmbStatusAcc.AddItem "LPD 3"
    CmbStatusAcc.AddItem "LPD 3<"
    
    cmdsql = "select * from contacteddesc where status='1' and jenis is not null "
    cmdsql = cmdsql & " and  jenis<>'CO-' order by jenis asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            CmbStatusAcc.AddItem IIf(IsNull(M_Objrs("jenis")), "", M_Objrs("jenis"))
            M_Objrs.MoveNext
        Wend
    End If
    CmbStatusAcc.AddItem "PTP-"
    Set M_Objrs = Nothing
End Sub

'@@21022013 Tambahan buat program filter account yang bisa di collect bersama
Private Sub IsiAgentCollectBersama()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim M_Objrs_TL As ADODB.Recordset
    
    CmbAgentCollBersama.clear
    
    cmdsql = "select distinct agent_asli from mgm "
    cmdsql = cmdsql + " where agent_asli is not null and agent='AKSESALL' "
    cmdsql = cmdsql + " and agent in (select userid from usertbl where usertype='1') "
    cmdsql = cmdsql + " order by agent_asli asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, admcdtext
    If M_Objrs.RecordCount > 0 Then
        CmbAgentCollBersama.AddItem "ALL"
        While Not M_Objrs.EOF
            CmbAgentCollBersama.AddItem M_Objrs("agent_asli")
            M_Objrs.MoveNext
        Wend
    
        'Load TLnya juga buat grouping
        cmdsql = "select userid from usertbl where usertype='6' order by userid asc "
        Set M_Objrs_TL = New ADODB.Recordset
        M_Objrs_TL.CursorLocation = adUseClient
        M_Objrs_TL.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_TL.RecordCount > 0 Then
            While Not M_Objrs_TL.EOF
                CmbAgentCollBersama.AddItem M_Objrs_TL("userid")
                M_Objrs_TL.MoveNext
            Wend
        End If
        Set M_Objrs_TL = Nothing
    End If
    Set M_Objrs = Nothing
    
End Sub


Private Sub IsiStatusCollectBersama()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbStatusCollBersama.clear
    
    cmdsql = "select distinct f_cek_new from mgm where agent='AKSESALL' "
    cmdsql = cmdsql + " and f_cek_new is not null order by f_cek_new asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        CmbStatusCollBersama.AddItem "ALL"
        CmbStatusCollBersama.AddItem "LPD 1"
        CmbStatusCollBersama.AddItem "LPD 2"
        CmbStatusCollBersama.AddItem "LPD 3"
        CmbStatusCollBersama.AddItem "LPD 3<"
        While Not M_Objrs.EOF
            CmbStatusCollBersama.AddItem Trim(UCase(Mid(M_Objrs("f_cek_new"), 1, 3)))
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub



