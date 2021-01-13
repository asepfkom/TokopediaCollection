VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmListRequestPTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Request PTP"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13245
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   11986
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   8421504
      TabCaption(0)   =   "List Request PTP"
      TabPicture(0)   =   "FrmListRequestPTP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Shape1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TxtLPAPayment"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LvPTP"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtJml"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmdCekall"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CmdUnCekAll"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CmdApprove"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CmdReject"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmdRefresh"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtLPDPayment"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "PB1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtCustid"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtNama"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CmdCari"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CmbTampilkan"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CmdApproveByPTP"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CmbApprove"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CmdApproveVP"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CmdExport"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "PanelExport"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "CD_save"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmd_SID"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "SEND PTP REJECTED"
      TabPicture(1)   =   "FrmListRequestPTP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(2)=   "Line2"
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(5)=   "LvPTPRejected"
      Tab(1).Control(6)=   "TxtJmlDataRejected"
      Tab(1).Control(7)=   "CmbJenisRejected"
      Tab(1).Control(8)=   "CmdCariRejected"
      Tab(1).Control(9)=   "TxtNamaRejected"
      Tab(1).Control(10)=   "TxtCustidRejected"
      Tab(1).Control(11)=   "CmbKembalikan"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "SEND PTP APPROVED"
      TabPicture(2)   =   "FrmListRequestPTP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command8"
      Tab(2).Control(1)=   "cmdsudahemail"
      Tab(2).Control(2)=   "cmrefresh"
      Tab(2).Control(3)=   "cmdsearch"
      Tab(2).Control(4)=   "txtsearch"
      Tab(2).Control(5)=   "cbsearch"
      Tab(2).Control(6)=   "TxtJmlApproved"
      Tab(2).Control(7)=   "LvPTPApproved"
      Tab(2).Control(8)=   "date1"
      Tab(2).Control(9)=   "date2"
      Tab(2).Control(10)=   "Shape2"
      Tab(2).Control(11)=   "Label2(2)"
      Tab(2).Control(12)=   "Label17"
      Tab(2).Control(13)=   "Label16"
      Tab(2).Control(14)=   "Label10"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "Approve By Pak Hamanto"
      TabPicture(3)   =   "FrmListRequestPTP.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label11"
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(2)=   "Label13"
      Tab(3).Control(3)=   "Label14"
      Tab(3).Control(4)=   "TxtTglApprove"
      Tab(3).Control(5)=   "PB2"
      Tab(3).Control(6)=   "LvHamanto"
      Tab(3).Control(7)=   "CmdCariAppHamanto"
      Tab(3).Control(8)=   "TxtCariNamaHamanto"
      Tab(3).Control(9)=   "TxtCustidHamanto"
      Tab(3).Control(10)=   "TxtJmlhAppHamanto"
      Tab(3).Control(11)=   "CmdApproveHamanto"
      Tab(3).Control(12)=   "CmdCekAllHamanto"
      Tab(3).Control(13)=   "CmdUnCekAllHamanto"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Change PTP Approval"
      TabPicture(4)   =   "FrmListRequestPTP.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label18"
      Tab(4).Control(1)=   "ListView1"
      Tab(4).Control(2)=   "Command3"
      Tab(4).Control(3)=   "Command4"
      Tab(4).Control(4)=   "Command5"
      Tab(4).Control(5)=   "Check1"
      Tab(4).Control(6)=   "Text1"
      Tab(4).Control(7)=   "Command6"
      Tab(4).Control(8)=   "Command7"
      Tab(4).ControlCount=   9
      Begin VB.CommandButton Command8 
         Caption         =   "Export"
         Height          =   375
         Left            =   -74880
         TabIndex        =   80
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Export"
         Height          =   375
         Left            =   -63480
         TabIndex        =   79
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Hst"
         Height          =   375
         Left            =   -63480
         TabIndex        =   78
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check All"
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reject"
         Height          =   495
         Left            =   -63360
         TabIndex        =   74
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Approve"
         Height          =   495
         Left            =   -63360
         TabIndex        =   73
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "View"
         Height          =   375
         Left            =   -63480
         TabIndex        =   72
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdsudahemail 
         Caption         =   "Already email"
         Height          =   375
         Left            =   -63120
         TabIndex        =   69
         Top             =   5700
         Width           =   1095
      End
      Begin VB.CommandButton cmrefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   -63000
         TabIndex        =   68
         Top             =   6300
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   -64080
         TabIndex        =   67
         Top             =   6300
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtsearch 
         Height          =   285
         Left            =   -69480
         TabIndex        =   62
         Top             =   5760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cbsearch 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":008C
         Left            =   -71280
         List            =   "FrmListRequestPTP.frx":009C
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   5760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmd_SID 
         Caption         =   "Add To SID"
         Height          =   375
         Left            =   11160
         TabIndex        =   59
         Top             =   4140
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   480
         Top             =   540
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel PanelExport 
         Height          =   1335
         Left            =   7680
         TabIndex        =   52
         Top             =   1980
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2355
         _Version        =   196610
         ActiveColors    =   -1  'True
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Txtlocation 
            Height          =   285
            Left            =   1800
            TabIndex        =   57
            Top             =   1080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
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
            Left            =   2280
            TabIndex        =   56
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Export"
            Height          =   375
            Left            =   1800
            TabIndex        =   55
            Top             =   480
            Width           =   735
         End
         Begin TDBDate6Ctl.TDBDate TdbDateExport 
            Height          =   285
            Left            =   240
            TabIndex        =   54
            Top             =   600
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   494
            Calendar        =   "FrmListRequestPTP.frx":00DF
            Caption         =   "FrmListRequestPTP.frx":01F7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmListRequestPTP.frx":0263
            Keys            =   "FrmListRequestPTP.frx":0281
            Spin            =   "FrmListRequestPTP.frx":02DF
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
         Begin VB.Label Label15 
            Caption         =   "Tanggal Approve"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export to Excel"
         Height          =   735
         Left            =   12360
         TabIndex        =   51
         Top             =   1980
         Width           =   735
      End
      Begin VB.CommandButton CmdUnCekAllHamanto 
         Caption         =   "&UnCek All"
         Height          =   315
         Left            =   -67080
         TabIndex        =   49
         Top             =   720
         Width           =   1395
      End
      Begin VB.CommandButton CmdCekAllHamanto 
         Caption         =   "&Cek All"
         Height          =   315
         Left            =   -68460
         TabIndex        =   50
         Top             =   720
         Width           =   1395
      End
      Begin VB.CommandButton CmdApproveHamanto 
         Caption         =   "&Approve"
         Height          =   435
         Left            =   -63720
         TabIndex        =   46
         Top             =   1140
         Width           =   1755
      End
      Begin VB.TextBox TxtJmlhAppHamanto 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73860
         TabIndex        =   44
         Text            =   "0"
         Top             =   6300
         Width           =   1215
      End
      Begin VB.TextBox TxtCustidHamanto 
         Height          =   285
         Left            =   -74160
         TabIndex        =   39
         Top             =   780
         Width           =   2115
      End
      Begin VB.TextBox TxtCariNamaHamanto 
         Height          =   285
         Left            =   -71280
         TabIndex        =   38
         Top             =   780
         Width           =   1635
      End
      Begin VB.CommandButton CmdCariAppHamanto 
         Caption         =   "&Cari"
         Height          =   315
         Left            =   -69600
         TabIndex        =   37
         Top             =   720
         Width           =   915
      End
      Begin VB.CommandButton CmdApproveVP 
         Caption         =   "To Be Approve By Pak Hamanto"
         Height          =   615
         Left            =   11160
         TabIndex        =   36
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox TxtJmlApproved 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73860
         TabIndex        =   34
         Text            =   "0"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton CmbKembalikan 
         Caption         =   "&Kembalikan ke list request PTP"
         Height          =   435
         Left            =   -65040
         TabIndex        =   32
         Top             =   1020
         Width           =   2955
      End
      Begin VB.TextBox TxtCustidRejected 
         Height          =   285
         Left            =   -74220
         TabIndex        =   28
         Top             =   1200
         Width           =   2115
      End
      Begin VB.TextBox TxtNamaRejected 
         Height          =   285
         Left            =   -71340
         TabIndex        =   27
         Top             =   1200
         Width           =   1635
      End
      Begin VB.CommandButton CmdCariRejected 
         Caption         =   "&Cari"
         Height          =   315
         Left            =   -69660
         TabIndex        =   26
         Top             =   1140
         Width           =   915
      End
      Begin VB.ComboBox CmbJenisRejected 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":0307
         Left            =   -67080
         List            =   "FrmListRequestPTP.frx":0311
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox TxtJmlDataRejected 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73860
         TabIndex        =   23
         Text            =   "0"
         Top             =   6240
         Width           =   1215
      End
      Begin VB.ComboBox CmbApprove 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":032E
         Left            =   11160
         List            =   "FrmListRequestPTP.frx":0344
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3660
         Width           =   1875
      End
      Begin VB.CommandButton CmdApproveByPTP 
         Caption         =   "&Approve PTP DISC. By SPV"
         Height          =   795
         Left            =   11160
         TabIndex        =   19
         Top             =   5340
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox CmbTampilkan 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":0395
         Left            =   7800
         List            =   "FrmListRequestPTP.frx":039F
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   3315
      End
      Begin VB.CommandButton CmdCari 
         Caption         =   "&Cari"
         Height          =   315
         Left            =   5340
         TabIndex        =   16
         Top             =   1080
         Width           =   915
      End
      Begin VB.TextBox TxtNama 
         Height          =   285
         Left            =   3660
         TabIndex        =   15
         Top             =   1140
         Width           =   1635
      End
      Begin VB.TextBox TxtCustid 
         Height          =   285
         Left            =   780
         TabIndex        =   13
         Top             =   1140
         Width           =   2115
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   6240
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox TxtLPDPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5460
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   11280
         TabIndex        =   8
         Top             =   2880
         Width           =   1395
      End
      Begin VB.CommandButton CmdReject 
         Caption         =   "Reject"
         Height          =   375
         Left            =   11280
         TabIndex        =   7
         Top             =   2340
         Width           =   1035
      End
      Begin VB.CommandButton CmdApprove 
         Caption         =   "&Approve"
         Height          =   375
         Left            =   11280
         TabIndex        =   6
         Top             =   1980
         Width           =   1035
      End
      Begin VB.CommandButton CmdUnCekAll 
         Caption         =   "&UnCek All"
         Height          =   375
         Left            =   11280
         TabIndex        =   5
         Top             =   1500
         Width           =   1395
      End
      Begin VB.CommandButton CmdCekall 
         Caption         =   "&Cek All"
         Height          =   375
         Left            =   11280
         TabIndex        =   4
         Top             =   1140
         Width           =   1395
      End
      Begin VB.TextBox TxtJml 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Text            =   "0"
         Top             =   6240
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   4620
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   1500
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   8149
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
      Begin TDBNumber6Ctl.TDBNumber TxtLPAPayment 
         Height          =   255
         Left            =   7740
         TabIndex        =   10
         Top             =   4680
         Visible         =   0   'False
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   450
         Calculator      =   "FrmListRequestPTP.frx":03BC
         Caption         =   "FrmListRequestPTP.frx":03DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":0448
         Keys            =   "FrmListRequestPTP.frx":0466
         Spin            =   "FrmListRequestPTP.frx":04B0
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   10147522
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin MSComctlLib.ListView LvPTPRejected 
         Height          =   4620
         Left            =   -74880
         TabIndex        =   22
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   1620
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   8149
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   8438015
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
      Begin MSComctlLib.ListView LvPTPApproved 
         Height          =   5160
         Left            =   -74940
         TabIndex        =   33
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   420
         Width           =   12960
         _ExtentX        =   22860
         _ExtentY        =   9102
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   8454016
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
      Begin MSComctlLib.ListView LvHamanto 
         Height          =   5040
         Left            =   -74880
         TabIndex        =   40
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   1140
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   8890
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
      Begin MSComctlLib.ProgressBar PB2 
         Height          =   315
         Left            =   -72480
         TabIndex        =   43
         Top             =   6300
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin TDBDate6Ctl.TDBDate TxtTglApprove 
         Height          =   285
         Left            =   -63600
         TabIndex        =   47
         Top             =   1860
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "FrmListRequestPTP.frx":04D8
         Caption         =   "FrmListRequestPTP.frx":05F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":065C
         Keys            =   "FrmListRequestPTP.frx":067A
         Spin            =   "FrmListRequestPTP.frx":06D8
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
      Begin TDBDate6Ctl.TDBDate date1 
         Height          =   285
         Left            =   -69480
         TabIndex        =   64
         Top             =   5760
         Visible         =   0   'False
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   503
         Calendar        =   "FrmListRequestPTP.frx":0700
         Caption         =   "FrmListRequestPTP.frx":0818
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":0884
         Keys            =   "FrmListRequestPTP.frx":08A2
         Spin            =   "FrmListRequestPTP.frx":0900
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
      Begin TDBDate6Ctl.TDBDate date2 
         Height          =   285
         Left            =   -67560
         TabIndex        =   66
         Top             =   5760
         Visible         =   0   'False
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "FrmListRequestPTP.frx":0928
         Caption         =   "FrmListRequestPTP.frx":0A40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":0AAC
         Keys            =   "FrmListRequestPTP.frx":0ACA
         Spin            =   "FrmListRequestPTP.frx":0B28
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   5460
         Left            =   -74880
         TabIndex        =   71
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   600
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   9631
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
      Begin VB.Label Label18 
         Caption         =   "Jumlah Amount"
         Height          =   495
         Left            =   -72480
         TabIndex        =   77
         Top             =   6180
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -65520
         Top             =   5820
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Belum di Email"
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
         Index           =   2
         Left            =   -65160
         TabIndex        =   70
         Top             =   5820
         Width           =   1275
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   -67920
         TabIndex        =   63
         Top             =   5760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "Search By :"
         Height          =   255
         Left            =   -72240
         TabIndex        =   61
         Top             =   5760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Belum di print"
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
         Index           =   1
         Left            =   11640
         TabIndex        =   58
         Top             =   6300
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   11280
         Top             =   6300
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Tanggal Approve:"
         Height          =   195
         Left            =   -63720
         TabIndex        =   48
         Top             =   1620
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   -74880
         TabIndex        =   45
         Top             =   6300
         Width           =   1035
      End
      Begin VB.Label Label12 
         Caption         =   "Custid:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   42
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Nama:"
         Height          =   195
         Left            =   -71940
         TabIndex        =   41
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   -74880
         TabIndex        =   35
         Top             =   5760
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Custid:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   31
         Top             =   1260
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "Nama:"
         Height          =   195
         Left            =   -72000
         TabIndex        =   30
         Top             =   1260
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   -68580
         X2              =   -68580
         Y1              =   1020
         Y2              =   1560
      End
      Begin VB.Label Label7 
         Caption         =   "Tampilkan hanya:"
         Height          =   195
         Left            =   -68460
         TabIndex        =   29
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label6 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   -74880
         TabIndex        =   24
         Top             =   6240
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Approve By:"
         Height          =   195
         Left            =   11520
         TabIndex        =   20
         Top             =   3420
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Tampilkan hanya:"
         Height          =   195
         Left            =   6420
         TabIndex        =   17
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   6360
         X2              =   6360
         Y1              =   960
         Y2              =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Nama:"
         Height          =   195
         Left            =   3000
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Custid:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   6240
         Width           =   1035
      End
   End
   Begin TDBDate6Ctl.TDBDate TDBDate3 
      Height          =   285
      Left            =   11400
      TabIndex        =   65
      Top             =   0
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "FrmListRequestPTP.frx":0B50
      Caption         =   "FrmListRequestPTP.frx":0C68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmListRequestPTP.frx":0CD4
      Keys            =   "FrmListRequestPTP.frx":0CF2
      Spin            =   "FrmListRequestPTP.frx":0D50
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
   Begin VB.Menu OH 
      Caption         =   "Ontario Hutagalung"
   End
End
Attribute VB_Name = "FrmListRequestPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StatusPTP As String
Dim PaymentTenor As Double
Dim jmlamount As String


Private Sub HeaderLog()
    LvPTP.ColumnHeaders.clear
    With LvPTP.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Nama CH", 3000
        .ADD 5, , "Status", 2000
        .ADD 6, , "Tanggal Approve", 2000
        .ADD 7, , "Tgl.Payment Effective", 2500
        .ADD 8, , "Total Amount", 1000
        .ADD 9, , "Tenor", 700
        .ADD 10, , "Pembayaran Via", 2000
        .ADD 11, , "Tgl.Tagih", 1500
        .ADD 12, , "Principal", 1000
        .ADD 13, , "Balance", 1000
        .ADD 14, , "Pembayaran Awal", 2000
        .ADD 15, , "Principal", 2000
        .ADD 16, , "Total Payment", 2000
        .ADD 17, , "Down Payment", 2000
        .ADD 18, , "Charge", 2000
        .ADD 19, , "Discount", 2000
        .ADD 20, , "From o/s balance %", 2000
        .ADD 21, , "Principal %", 2000
        .ADD 22, , "Justtification", 2000
        .ADD 23, , "Fax", 800
        .ADD 24, , "When Talking Surlun", 800
        .ADD 25, , "KTP", 800
        .ADD 26, , "Surper", 800
        .ADD 27, , "Billing", 800
        .ADD 28, , "Other", 800
        .ADD 29, , "Agent", 800
        .ADD 30, , "DOB", 1000
        .ADD 31, , "Ket.Other", 1000
        
        '@@ 16-07-2012 Tambahan Payment Handle
        .ADD 32, , "Payment Handle", 2000
        
        '@@17-07-2012 Tambahan Occupation dan Reason
        .ADD 33, , "Occupation", 2000
        .ADD 34, , "Reason", 2000
        
    End With
End Sub


Private Sub HeaderLogRejected()
    LvPTPRejected.ColumnHeaders.clear
    With LvPTPRejected.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Nama CH", 3000
        .ADD 5, , "Status", 2000
        .ADD 6, , "Tanggal Approve", 2000
        .ADD 7, , "Tgl.Payment Effective", 2500
        .ADD 8, , "Total Amount", 1000
        .ADD 9, , "Tenor", 700
        .ADD 10, , "Pembayaran Via", 2000
        .ADD 11, , "Tgl.Tagih", 1500
        .ADD 12, , "Principal", 1000
        .ADD 13, , "Balance", 1000
        .ADD 14, , "Pembayaran Awal", 2000
        .ADD 15, , "Principal", 2000
        .ADD 16, , "Total Payment", 2000
        .ADD 17, , "Down Payment", 2000
        .ADD 18, , "Charge", 2000
        .ADD 19, , "Discount", 2000
        .ADD 20, , "From o/s balance %", 2000
        .ADD 21, , "Principal %", 2000
        .ADD 22, , "Justtification", 2000
        .ADD 23, , "Fax", 800
        .ADD 24, , "When Talking Surlun", 800
        .ADD 25, , "KTP", 800
        .ADD 26, , "Surper", 800
        .ADD 27, , "Billing", 800
        .ADD 28, , "Other", 800
        .ADD 29, , "Agent", 800
        .ADD 30, , "DOB", 1000
        .ADD 31, , "Ket.Other", 1000
    End With
End Sub

Private Sub HeaderLogApproved()
'    LvPTPApproved.ColumnHeaders.CLEAR
'    With LvPTPApproved.ColumnHeaders
'        .ADD 1, , "ID", 500
'        .ADD 2, , "Jenis PTP", 1000
'        .ADD 3, , "Custid", 2000
'        .ADD 4, , "Nama CH", 3000
'        .ADD 5, , "Status", 2000
'        .ADD 6, , "Tanggal Approve", 2000
'        .ADD 7, , "Tgl.Payment Effective", 2500
'        .ADD 8, , "Total Amount", 1000
'        .ADD 9, , "Tenor", 700
'        .ADD 10, , "Pembayaran Via", 2000
'        .ADD 11, , "Tgl.Tagih", 1500
'        .ADD 12, , "Principal", 1000
'        .ADD 13, , "Balance", 1000
'        .ADD 14, , "Pembayaran Awal", 2000
'        .ADD 15, , "Principal", 2000
'        .ADD 16, , "Total Payment", 2000
'        .ADD 17, , "Down Payment", 2000
'        .ADD 18, , "Charge", 2000
'        .ADD 19, , "Discount", 2000
'        .ADD 20, , "From o/s balance %", 2000
'        .ADD 21, , "Principal %", 2000
'        .ADD 22, , "Justtification", 2000
'        .ADD 23, , "Fax", 800
'        .ADD 24, , "When Talking Surlun", 800
'        .ADD 25, , "KTP", 800
'        .ADD 26, , "Surper", 800
'        .ADD 27, , "Billing", 800
'        .ADD 28, , "Other", 800
'        .ADD 29, , "Agent", 800
'        .ADD 30, , "DOB", 1000
'        .ADD 31, , "Ket.Other", 1000
'        .ADD 32, , "Level Approval", 2000
'    End With

'jejaktian
    LvPTPApproved.ColumnHeaders.clear
'    With LvPTPApproved.ColumnHeaders
'        .ADD 1, , "ID", 500
'        .ADD 2, , "Jenis PTP", 1000
'        .ADD 3, , "Custid", 2000
'        .ADD 4, , "Nama CH", 3000
'        .ADD 5, , "Status", 2000
'        .ADD 6, , "Tanggal Approve", 2000
'        .ADD 7, , "Level Approval", 2000
'        .ADD 8, , "Tanggal Request by Email", 2000
'    End With
'12feb2018
    With LvPTPApproved.ColumnHeaders
        .ADD 1, , "CREATE DATE", 500
        .ADD 2, , "PRODUCT", 1000
        .ADD 3, , "REGION", 2000
        .ADD 4, , "CUSTID", 3000
        .ADD 5, , "NAME", 2000
        .ADD 6, , "WO DATE", 2000
        .ADD 7, , "PRINCIPAL", 2000
        .ADD 8, , "BALANCE", 2000
        .ADD 9, , "TOTALPAY", 2000
        .ADD 10, , "DP", 2000
        .ADD 11, , "TENOR", 2000
        .ADD 12, , "DISCOUNTAMOUNT", 2000
        .ADD 13, , "PRINCIPAL%", 2000
        .ADD 14, , "BALANCE%", 2000
        .ADD 15, , "SEGMENT, 2000"
        .ADD 16, , "JUSTIFICATION", 2000
        .ADD 17, , "JML", 2000
    End With

End Sub

Private Sub isilistappptp()
    'qusel = "select * from tblnegoptp_temp_app order by id asc"
    'qusel = "select a.*,b.team,c.lastpay,c.pay_dt from tblnegoptp_temp_app a,usertbl b,mgm c where a.agent = b.userid and a.custid = c.custid order by id asc "
    
    jmlamount = "0"
    
'    qusel = " select a.*,c.payment,c.paydate from ("
'    qusel = qusel + vbCrLf + " select a.*,b.team from tblnegoptp_temp_app a,usertbl b where a.agent = b.userid"
'    qusel = qusel + vbCrLf + " ) a left join"
'    qusel = qusel + vbCrLf + " ("
'    qusel = qusel + vbCrLf + " select * from tbllunas where to_char(paydate,'yyyy-mm') = to_char(now(),'yyyy-mm')"
'    qusel = qusel + vbCrLf + " ) c on a.custid = c.custid"
'    qusel = qusel + vbCrLf + " order by a.id asc"
    
    qusel = " select a.*, mgm.name from ("
    qusel = qusel + vbCrLf + " select a.*,c.payment,c.paydate from ("
    qusel = qusel + vbCrLf + " select a.*,b.team from tblnegoptp_temp_app a,usertbl b where a.agent = b.userid"
    qusel = qusel + vbCrLf + " ) a left join"
    qusel = qusel + vbCrLf + " ("
    qusel = qusel + vbCrLf + " select * from tbllunas where to_char(paydate,'yyyy-mm') = to_char(now(),'yyyy-mm')"
    qusel = qusel + vbCrLf + " ) c on a.custid = c.custid"
    qusel = qusel + vbCrLf + " order by a.id asc"
    qusel = qusel + vbCrLf + " ) a left join mgm on a.custid = mgm.custid"

    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open qusel, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ListView1.ListItems.clear
    ListView1.Checkboxes = True
    Check1.Caption = "Check All " & M_Objrs.RecordCount
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            Set listItem = ListView1.ListItems.ADD(, , M_Objrs("custid"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("id")), "", M_Objrs("id"))
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("name")), "", M_Objrs("name"))
                listItem.SubItems(3) = Format(IIf(IsNull(M_Objrs("promisedate")), "", M_Objrs("promisedate")), "yyyy-mm-dd")
                listItem.SubItems(4) = IIf(IsNull(M_Objrs("promisepay")), "", M_Objrs("promisepay"))
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("statusptp")), "", M_Objrs("statusptp"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("via")), "", M_Objrs("via"))
                listItem.SubItems(7) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                listItem.SubItems(8) = IIf(IsNull(M_Objrs("team")), "", M_Objrs("team"))
                listItem.SubItems(9) = Format(IIf(IsNull(M_Objrs("reqdate")), "", M_Objrs("reqdate")), "yyyy-mm-dd")
                listItem.SubItems(10) = Format(IIf(IsNull(M_Objrs("paydate")), "", M_Objrs("paydate")), "yyyy-mm-dd")
                listItem.SubItems(11) = IIf(IsNull(M_Objrs("payment")), "", M_Objrs("payment"))
                listItem.SubItems(12) = Format(IIf(IsNull(M_Objrs("tgltagih")), "", M_Objrs("tgltagih")), "yyyy-mm-dd")
                listItem.SubItems(13) = IIf(IsNull(M_Objrs("tenor")), "", M_Objrs("tenor"))
                
                jmlamount = jmlamount + IIf(IsNull(M_Objrs("promisepay")), "", M_Objrs("promisepay"))
                
                If IIf(IsNull(M_Objrs("promisepay")), "", M_Objrs("promisepay")) > IIf(IsNull(M_Objrs("payment")), "", M_Objrs("payment")) Then
                    For K = 1 To 12
                        listItem.ListSubItems(K).ForeColor = vbRed
                        listItem.ListSubItems(K).Bold = True
                    Next K
                End If
                
                M_Objrs.MoveNext
        Wend
        Text1.text = Format(jmlamount, "###,##")
    Else
        MsgBox "Permintaan Change PTP Kosong"
    End If
    
End Sub

Private Sub isilisthstchange()
    qusel = "select * from tblnegoptp_temp_app_log order by id desc limit 1000"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open qusel, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ListView1.ListItems.clear
    
    ListView1.Checkboxes = False
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            Set listItem = ListView1.ListItems.ADD(, , M_Objrs("custid"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("id")), "", M_Objrs("id"))
                listItem.SubItems(2) = Format(IIf(IsNull(M_Objrs("promisedate")), "", M_Objrs("promisedate")), "yyyy-mm-dd")
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("promisepay")), "", M_Objrs("promisepay"))
                listItem.SubItems(4) = IIf(IsNull(M_Objrs("statusptp")), "", M_Objrs("statusptp"))
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("via")), "", M_Objrs("via"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                listItem.SubItems(7) = Format(IIf(IsNull(M_Objrs("reqdate")), "", M_Objrs("reqdate")), "yyyy-mm-dd")
                listItem.SubItems(8) = Format(IIf(IsNull(M_Objrs("tgltagih")), "", M_Objrs("tgltagih")), "yyyy-mm-dd")
                listItem.SubItems(9) = IIf(IsNull(M_Objrs("tenor")), "", M_Objrs("tenor"))
                listItem.SubItems(10) = IIf(IsNull(M_Objrs("appby")), "", M_Objrs("appby"))
                listItem.SubItems(11) = Format(IIf(IsNull(M_Objrs("appdate")), "", M_Objrs("appdate")), "yyyy-mm-dd")
                listItem.SubItems(12) = IIf(IsNull(M_Objrs("status")), "", M_Objrs("status"))
      
                M_Objrs.MoveNext
        Wend
    End If
    
End Sub


Public Sub isilog()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
    Call HeaderLog
    '<9feb2018
    cmdsql = "select * from tblsendptp where id is not null "
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        cmdsql = cmdsql + " and agent in "
        cmdsql = cmdsql + "(select userid from usertbl where team='"
        cmdsql = cmdsql + MDIForm1.Text1.text + "' and usertype='1' and aktif='0') "
        
        'CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
    End If
    
    If Left(MDIForm1.Text2.text, 2) = "AM" Then
        cmdsql = cmdsql + " and agent in "
        cmdsql = cmdsql + " (select userid from usertbl where team in (select tl from tblsettingam  where am = '" + MDIForm1.Text1.text + "') and usertype='1' and aktif = '0' ) "
    End If
    If TxtNama.text <> Empty Then
        cmdsql = cmdsql + " and vcustname like '%"
        cmdsql = cmdsql + TxtNama.text + "%' "
    End If
    If TxtCustid.text <> Empty Then
        cmdsql = cmdsql + " and custid like '%"
        cmdsql = cmdsql + TxtCustid.text + "%' "
    End If
    
    If CmbTampilkan.text = "PTP NO DISC." Then
       cmdsql = cmdsql + " and jenis_ptp='PTP No Discount' "
       cmdsql = cmdsql + " and status='0' "
    End If
    If CmbTampilkan.text = "PTP DISC." Then
        cmdsql = cmdsql + " and jenis_ptp='PTP Discount' "
        cmdsql = cmdsql + " and status='0'"
    End If
    If CmbTampilkan.text = "PTP DISC. APPROVED" Then
        cmdsql = cmdsql + " and jenis_ptp='PTP Discount' "
        cmdsql = cmdsql + " and status='1' "
    End If
    
    
   '@@221012 Tambahan buat approve by VP
    cmdsql = cmdsql + " and  sts_app_vp is null "
    
    cmdsql = cmdsql + " order by tgldata desc"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPTP.ListItems.clear
    txtjml.text = M_Objrs.RecordCount
    
    If M_Objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        While Not M_Objrs.EOF
            'On Error Resume Next
            Set listItem = LvPTP.ListItems.ADD(, , M_Objrs("id"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("vcustname")), "", M_Objrs("vcustname"))
                
                If M_Objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_Objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_Objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                listItem.SubItems(4) = STATUS
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("tgl_approve")), "", Format(M_Objrs("tgl_approve"), "yyyy-mm-dd"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("date_payment_effective")), "", Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd"))
                listItem.SubItems(7) = IIf(IsNull(M_Objrs("total_amount_deal")), "0", M_Objrs("total_amount_deal"))
                listItem.SubItems(8) = IIf(IsNull(M_Objrs("tenor")), "1", M_Objrs("tenor"))
                listItem.SubItems(9) = IIf(IsNull(M_Objrs("pembayaran_via")), "", M_Objrs("pembayaran_via"))
                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tgl_tagih")), "", Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd"))
                listItem.SubItems(11) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
                listItem.SubItems(12) = IIf(IsNull(M_Objrs("balance")), "0", M_Objrs("balance"))
                listItem.SubItems(13) = IIf(IsNull(M_Objrs("pembayaran_awal")), "0", M_Objrs("pembayaran_awal"))
                listItem.SubItems(14) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
                listItem.SubItems(15) = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
                listItem.SubItems(16) = IIf(IsNull(M_Objrs("ndownpay")), "0", M_Objrs("ndownpay"))
                listItem.SubItems(17) = IIf(IsNull(M_Objrs("ncharge")), "0", M_Objrs("ncharge"))
                listItem.SubItems(18) = IIf(IsNull(M_Objrs("ndiscountamt")), "0", M_Objrs("ndiscountamt"))
                listItem.SubItems(19) = IIf(IsNull(M_Objrs("vosbalance")), "", M_Objrs("vosbalance"))
                listItem.SubItems(20) = IIf(IsNull(M_Objrs("vosprincipal")), "", M_Objrs("vosprincipal"))
                listItem.SubItems(21) = IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust"))
                listItem.SubItems(22) = IIf(IsNull(M_Objrs("chkfaxed")), "", M_Objrs("chkfaxed"))
                listItem.SubItems(23) = IIf(IsNull(M_Objrs("chkwentalking")), "", M_Objrs("chkwentalking"))
                listItem.SubItems(24) = IIf(IsNull(M_Objrs("chkktp")), "", M_Objrs("chkktp"))
                listItem.SubItems(25) = IIf(IsNull(M_Objrs("chksup")), "", M_Objrs("chksup"))
                listItem.SubItems(26) = IIf(IsNull(M_Objrs("chkbillings")), "", M_Objrs("chkbillings"))
                listItem.SubItems(27) = IIf(IsNull(M_Objrs("chkothers")), "", M_Objrs("chkothers"))
                listItem.SubItems(28) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                  
                If IsNull(M_Objrs("dob")) = True Or M_Objrs("dob") = "" Or M_Objrs("dob") = Empty Then
                    DOB = ""
                Else
                    DOB = Format(M_Objrs("dob"), "yyyy-mm-dd")
                End If
                 
                listItem.SubItems(29) = DOB
                listItem.SubItems(30) = IIf(IsNull(M_Objrs("ket_other")), "", M_Objrs("ket_other"))
                listItem.SubItems(31) = IIf(IsNull(M_Objrs("payment_handle")), "", M_Objrs("payment_handle"))
                
                listItem.SubItems(32) = IIf(IsNull(M_Objrs("occupation")), "", M_Objrs("occupation"))
                listItem.SubItems(33) = IIf(IsNull(M_Objrs("reason")), "", M_Objrs("reason"))
                'listItem.SubItems(34) = IIf(IsNull(M_Objrs("tgl_send_email")), "", M_Objrs("tgl_send_email"))

                ' Tandain klo ini belum di print 13 Okt 2014
                If IIf(IsNull(M_Objrs("s_print")), 0, M_Objrs("s_print")) = 0 Then
                    For K = 1 To 7
                        listItem.ListSubItems(K).ForeColor = vbRed
                    Next K
                End If
                ' ------------------------------------------
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub


Public Sub IsiLogRejected()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
    cmdsql = "select * from tblsendptp_log_reject where id is not null "
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        cmdsql = cmdsql + " and agent in "
        cmdsql = cmdsql + "(select userid from usertbl where team='"
        cmdsql = cmdsql + MDIForm1.Text1.text + "' and usertype='1' and aktif='0') "
        
        'CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
    End If
    If TxtNamaRejected.text <> Empty Then
        cmdsql = cmdsql + " and vcustname like '%"
        cmdsql = cmdsql + TxtNamaRejected.text + "%' "
    End If
    If TxtCustidRejected.text <> Empty Then
        cmdsql = cmdsql + " and custid like '%"
        cmdsql = cmdsql + TxtCustidRejected.text + "%' "
    End If
    
    If CmbJenisRejected.text = "PTP NO DISC." Then
       cmdsql = cmdsql + " and jenis_ptp='PTP No Discount' "
       cmdsql = cmdsql + " and status='0' "
    End If
    If CmbJenisRejected.text = "PTP DISC." Then
        cmdsql = cmdsql + " and jenis_ptp='PTP Discount' "
        cmdsql = cmdsql + " and status='0'"
    End If
    If CmbJenisRejected.text = "PTP DISC. APPROVED" Then
        cmdsql = cmdsql + " and jenis_ptp='PTP Discount' "
        cmdsql = cmdsql + " and status='1' "
    End If
    
    
    
    cmdsql = cmdsql + " order by tgldata desc limit 300 "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPTPRejected.ListItems.clear
    TxtJmlDataRejected.text = M_Objrs.RecordCount
    
    If M_Objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        While Not M_Objrs.EOF
            'On Error Resume Next
            Set listItem = LvPTPRejected.ListItems.ADD(, , M_Objrs("id"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("vcustname")), "", M_Objrs("vcustname"))
                
                If M_Objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_Objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_Objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                listItem.SubItems(4) = STATUS
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("tgl_approve")), "", Format(M_Objrs("tgl_approve"), "yyyy-mm-dd"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("date_payment_effective")), "", Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd"))
                listItem.SubItems(7) = IIf(IsNull(M_Objrs("total_amount_deal")), "0", M_Objrs("total_amount_deal"))
                listItem.SubItems(8) = IIf(IsNull(M_Objrs("tenor")), "1", M_Objrs("tenor"))
                listItem.SubItems(9) = IIf(IsNull(M_Objrs("pembayaran_via")), "", M_Objrs("pembayaran_via"))
                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tgl_tagih")), "", Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd"))
                listItem.SubItems(11) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
                listItem.SubItems(12) = IIf(IsNull(M_Objrs("balance")), "0", M_Objrs("balance"))
                listItem.SubItems(13) = IIf(IsNull(M_Objrs("pembayaran_awal")), "0", M_Objrs("pembayaran_awal"))
                listItem.SubItems(14) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
                listItem.SubItems(15) = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
                listItem.SubItems(16) = IIf(IsNull(M_Objrs("ndownpay")), "0", M_Objrs("ndownpay"))
                listItem.SubItems(17) = IIf(IsNull(M_Objrs("ncharge")), "0", M_Objrs("ncharge"))
                listItem.SubItems(18) = IIf(IsNull(M_Objrs("ndiscountamt")), "0", M_Objrs("ndiscountamt"))
                listItem.SubItems(19) = IIf(IsNull(M_Objrs("vosbalance")), "", M_Objrs("vosbalance"))
                listItem.SubItems(20) = IIf(IsNull(M_Objrs("vosprincipal")), "", M_Objrs("vosprincipal"))
                listItem.SubItems(21) = IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust"))
                listItem.SubItems(22) = IIf(IsNull(M_Objrs("chkfaxed")), "", M_Objrs("chkfaxed"))
                listItem.SubItems(23) = IIf(IsNull(M_Objrs("chkwentalking")), "", M_Objrs("chkwentalking"))
                listItem.SubItems(24) = IIf(IsNull(M_Objrs("chkktp")), "", M_Objrs("chkktp"))
                listItem.SubItems(25) = IIf(IsNull(M_Objrs("chksup")), "", M_Objrs("chksup"))
                listItem.SubItems(26) = IIf(IsNull(M_Objrs("chkbillings")), "", M_Objrs("chkbillings"))
                listItem.SubItems(27) = IIf(IsNull(M_Objrs("chkothers")), "", M_Objrs("chkothers"))
                listItem.SubItems(28) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                  
                If IsNull(M_Objrs("dob")) = True Or M_Objrs("dob") = "" Or M_Objrs("dob") = Empty Then
                    DOB = ""
                Else
                    DOB = Format(M_Objrs("dob"), "yyyy-mm-dd")
                End If
                 
                listItem.SubItems(29) = DOB
                listItem.SubItems(30) = IIf(IsNull(M_Objrs("ket_other")), "", M_Objrs("ket_other"))
                'listItem.SubItems(31) = IIf(IsNull(M_Objrs("tgl_send_email")), "", M_Objrs("tgl_send_email"))
 
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub


Public Sub IsiLogApproved()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
'<9feb2018
    'cmdsql = "select * from tblsendptp_log_approve where id is not null "
'=======================================================================
        
    cmdsql = "select tgl_approve as create_date, acc_type as product, region, custid, vcustname as customername, B_D as wodate, principal, balance, total_amount_deal as totalpay, pembayaran_awal as dp, tenor, ndiscountamt as discountamount, vosprincipal as percenprincipal, vosbalance as percenbalance, segment, vjust as justification, jml from ("
    cmdsql = vbCrLf & cmdsql & " select tblsendptp_log_approve.*, a.segment, b.jml, a.acc_type, a.region, a.B_D from tblsendptp_log_approve left join (select custid,segment,acc_type,region,B_D from mgm where 1=1"
    
    
'    cmdsql = "select tblsendptp_log_approve.*, a.segment, b.jml from tblsendptp_log_approve left join (select custid,segment from mgm where segment <> '' and segment is not null) a on" & vbCrLf
'    cmdsql = cmdsql + " tblsendptp_log_approve.custid = a.custid left join (select * from (select custid, count(custid) as jml from tblsendptp_log_approve group by 1) a) b on tblsendptp_log_approve.custid = b.custid where id is not null "
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        cmdsql = cmdsql + " and agent in "
        cmdsql = cmdsql + "(select userid from usertbl where team='"
        cmdsql = cmdsql + MDIForm1.Text1.text + "' and usertype='1' and aktif='0') "
        
        'CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
    End If
    
    cmdsql = vbCrLf & cmdsql & " ) a on"
    cmdsql = vbCrLf & cmdsql & " tblsendptp_log_approve.custid = a.custid left join (select * from (select custid, count(custid) as jml from tblsendptp_log_approve group by 1) a) b on tblsendptp_log_approve.custid = b.custid where id is not null and jenis_ptp = 'PTP Discount' order by tgldata desc limit 300"
    cmdsql = vbCrLf & cmdsql & " ) abc"
    
    
'    If TxtNamaRejected.Text <> Empty Then
'        Cmdsql = Cmdsql + " and vcustname like '%"
'        Cmdsql = Cmdsql + TxtNamaRejected.Text + "%' "
'    End If
'    If TxtCustidRejected.Text <> Empty Then
'        Cmdsql = Cmdsql + " and custid like '%"
'        Cmdsql = Cmdsql + TxtCustidRejected.Text + "%' "
'    End If
'
'    If CmbJenisRejected.Text = "PTP NO DISC." Then
'       Cmdsql = Cmdsql + " and jenis_ptp='PTP No Discount' "
'       Cmdsql = Cmdsql + " and status='0' "
'    End If
'    If CmbJenisRejected.Text = "PTP DISC." Then
'        Cmdsql = Cmdsql + " and jenis_ptp='PTP Discount' "
'        Cmdsql = Cmdsql + " and status='0'"
'    End If
'    If CmbJenisRejected.Text = "PTP DISC. APPROVED" Then
'        Cmdsql = Cmdsql + " and jenis_ptp='PTP Discount' "
'        Cmdsql = Cmdsql + " and status='1' "
'    End If
    'jejaktian16032016
    
    
    '<12feb2018
    'cmdsql = cmdsql + "and jenis_ptp = 'PTP Discount'"
    'cmdsql = cmdsql + " order by tgldata desc limit 300 "
    '============================================
    
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPTPApproved.ListItems.clear
    TxtJmlApproved.text = M_Objrs.RecordCount
    
    If M_Objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        Dim discount As Double
        Dim S As String
        
        While Not M_Objrs.EOF
            'On Error Resume Next
'            Set listItem = LvPTPApproved.ListItems.ADD(, , M_Objrs("id"))
'                listItem.SubItems(1) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
'                listItem.SubItems(2) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
'                listItem.SubItems(3) = IIf(IsNull(M_Objrs("vcustname")), "", M_Objrs("vcustname"))
'
'                If M_Objrs("status") = "0" Then
'                    STATUS = "Belum di Approve"
'                End If
'                If M_Objrs("status") = "1" Then
'                    STATUS = "Approve"
'                End If
'                If M_Objrs("status") = "2" Then
'                    STATUS = "Rejected"
'                End If
'
'         discount = IIf(IsNull(M_Objrs("ndiscountamt")), "0", M_Objrs("ndiscountamt"))
'            If discount > 0 And discount <= 2000000 Then
'                S = "Coll SPV"
'            ElseIf discount <= 10000000 Then
'                S = "Coll Band 6"
'            ElseIf discount <= 20000000 Then
'                S = "Coll Band 5"
'            ElseIf discount <= 30000000 Then
'                S = "Coll Band 4"
'            ElseIf discount <= 50000000 Then
'                S = "Head of Coll"
'            ElseIf discount <= 100000000 Then
'                S = "Head of CCC"
'            ElseIf discount <= 2300000000# Then
'                S = "Head of CRM"
'            End If
'
'                listItem.SubItems(4) = STATUS
'                listItem.SubItems(5) = IIf(IsNull(M_Objrs("tgl_approve")), "", Format(M_Objrs("tgl_approve"), "yyyy-mm-dd"))
'                'jejaktianremark16032016
''                listItem.SubItems(6) = IIf(IsNull(M_Objrs("date_payment_effective")), "", Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd"))
''                listItem.SubItems(7) = IIf(IsNull(M_Objrs("total_amount_deal")), "0", M_Objrs("total_amount_deal"))
''                listItem.SubItems(8) = IIf(IsNull(M_Objrs("tenor")), "1", M_Objrs("tenor"))
''                listItem.SubItems(9) = IIf(IsNull(M_Objrs("pembayaran_via")), "", M_Objrs("pembayaran_via"))
''                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tgl_tagih")), "", Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd"))
''                listItem.SubItems(11) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
''                listItem.SubItems(12) = IIf(IsNull(M_Objrs("balance")), "0", M_Objrs("balance"))
''                listItem.SubItems(13) = IIf(IsNull(M_Objrs("pembayaran_awal")), "0", M_Objrs("pembayaran_awal"))
''                listItem.SubItems(14) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
''                listItem.SubItems(15) = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
''                listItem.SubItems(16) = IIf(IsNull(M_Objrs("ndownpay")), "0", M_Objrs("ndownpay"))
''                listItem.SubItems(17) = IIf(IsNull(M_Objrs("ncharge")), "0", M_Objrs("ncharge"))
''                listItem.SubItems(6) = discount
'
''                listItem.SubItems(19) = IIf(IsNull(M_Objrs("vosbalance")), "", M_Objrs("vosbalance"))
''                listItem.SubItems(20) = IIf(IsNull(M_Objrs("vosprincipal")), "", M_Objrs("vosprincipal"))
''                listItem.SubItems(21) = IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust"))
''                listItem.SubItems(22) = IIf(IsNull(M_Objrs("chkfaxed")), "", M_Objrs("chkfaxed"))
''                listItem.SubItems(23) = IIf(IsNull(M_Objrs("chkwentalking")), "", M_Objrs("chkwentalking"))
''                listItem.SubItems(24) = IIf(IsNull(M_Objrs("chkktp")), "", M_Objrs("chkktp"))
''                listItem.SubItems(25) = IIf(IsNull(M_Objrs("chksup")), "", M_Objrs("chksup"))
''                listItem.SubItems(26) = IIf(IsNull(M_Objrs("chkbillings")), "", M_Objrs("chkbillings"))
''                listItem.SubItems(27) = IIf(IsNull(M_Objrs("chkothers")), "", M_Objrs("chkothers"))
''                listItem.SubItems(28) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
''
''                If IsNull(M_Objrs("dob")) = True Or M_Objrs("dob") = "" Or M_Objrs("dob") = Empty Then
''                    DOB = ""
''                Else
''                    DOB = Format(M_Objrs("dob"), "yyyy-mm-dd")
''                End If
''
''                listItem.SubItems(29) = DOB
''                listItem.SubItems(30) = IIf(IsNull(M_Objrs("ket_other")), "", M_Objrs("ket_other"))
'                '==============================================
'                listItem.SubItems(6) = S
'                listItem.SubItems(7) = IIf(IsNull(M_Objrs("tgl_send_email")), "", Format(M_Objrs("tgl_send_email"), "yyyy-mm-dd"))
'
'
'                ' Tandain klo ini belum di send_email 'jejaktian 14032016
'                If IIf(IsNull(M_Objrs("tgl_send_email")), 0, M_Objrs("tgl_send_email")) = 0 Then
'                    For K = 1 To 7
'                        listItem.ListSubItems(K).ForeColor = vbRed
'                        listItem.ListSubItems(K).Bold = True
'                    Next K
'                End If
'                ' ------------------------------------------

            Set listItem = LvPTPApproved.ListItems.ADD(, , Format(M_Objrs("create_date"), "yyyy-mm-dd hh:nn:ss"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("product")), "", M_Objrs("product"))
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("region")), "", M_Objrs("region"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
                listItem.SubItems(4) = IIf(IsNull(M_Objrs("customername")), "", M_Objrs("customername"))
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("wodate")), "", M_Objrs("wodate"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("principal")), "", M_Objrs("principal"))
                listItem.SubItems(7) = IIf(IsNull(M_Objrs("balance")), "", M_Objrs("balance"))
                listItem.SubItems(8) = IIf(IsNull(M_Objrs("totalpay")), "", M_Objrs("totalpay"))
                listItem.SubItems(9) = IIf(IsNull(M_Objrs("dp")), "", M_Objrs("dp"))
                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tenor")), "", M_Objrs("tenor"))
                listItem.SubItems(11) = IIf(IsNull(M_Objrs("discountamount")), "", M_Objrs("discountamount"))
                listItem.SubItems(12) = IIf(IsNull(M_Objrs("percenprincipal")), "", M_Objrs("percenprincipal"))
                listItem.SubItems(13) = IIf(IsNull(M_Objrs("percenbalance")), "", M_Objrs("percenbalance"))
                listItem.SubItems(14) = IIf(IsNull(M_Objrs("segment")), "", M_Objrs("segment"))
                listItem.SubItems(15) = IIf(IsNull(M_Objrs("justification")), "", M_Objrs("justification"))
                listItem.SubItems(16) = IIf(IsNull(M_Objrs("jml")), "", M_Objrs("jml"))
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub


Private Sub cbsearch_Click()
Dim S As String
Dim M_Objrs As ADODB.Recordset
Dim listItem As listItem
Dim sql As String

    If cbsearch.text = "Tanggal Request by Email" Then
        S = "Tanggal Request by Email"
        txtsearch.Visible = False
        txtsearch.text = ""
        date1.Visible = True
        date2.Visible = True
        Label17.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        date1.text = ""
        date2.text = ""
    ElseIf cbsearch.text = "Cust ID" Then
        S = "CustId"
        txtsearch.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        txtsearch.text = ""
        date1.Visible = False
        date2.Visible = False
        Label17.Visible = False
    ElseIf cbsearch.text = "Nama Cust" Then
        S = "vcustname"
        txtsearch.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        txtsearch.text = ""
        date1.Visible = False
        date2.Visible = False
        Label17.Visible = False
    ElseIf cbsearch.text = "Tanggal Approve" Then
        S = "tgl_approve"
        txtsearch.Visible = False
        txtsearch.text = ""
        date1.Visible = True
        date2.Visible = True
        Label17.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        date1.text = ""
        date2.text = ""
    End If
       
'sql = "select * from tblsendptp_log_approve where id is not null "
'sql = sql + "and '" & s & "' =  '" & txtsearch.Text & "' or '" & s & "' between '" & Format(date1.Value, "yyyy/mm/dd") & "' and '" & Format(date2.Value, "yyyy/mm/dd") & "' order by tgldata desc limit 300 "
'M_OBJCONN.Execute sql
'
'LvPTPApproved.ListItems.CLEAR
'    TxtJmlApproved.Text = M_Objrs.RecordCount
'
'    If M_Objrs.RecordCount > 0 Then
'        Dim STATUS As String
'        Dim DOB As String
'        While Not M_Objrs.EOF
'            'On Error Resume Next
'            Set listItem = LvPTPApproved.ListItems.ADD(, , M_Objrs("id"))
'                listItem.SubItems(1) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
'                listItem.SubItems(2) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
'                listItem.SubItems(3) = IIf(IsNull(M_Objrs("vcustname")), "", M_Objrs("vcustname"))
'
'                If M_Objrs("status") = "0" Then
'                    STATUS = "Belum di Approve"
'                End If
'                If M_Objrs("status") = "1" Then
'                    STATUS = "Approve"
'                End If
'                If M_Objrs("status") = "2" Then
'                    STATUS = "Rejected"
'                End If
'
'                listItem.SubItems(4) = STATUS
'                listItem.SubItems(5) = IIf(IsNull(M_Objrs("tgl_approve")), "", Format(M_Objrs("tgl_approve"), "yyyy-mm-dd"))
'                listItem.SubItems(6) = IIf(IsNull(M_Objrs("date_payment_effective")), "", Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd"))
'                listItem.SubItems(7) = IIf(IsNull(M_Objrs("total_amount_deal")), "0", M_Objrs("total_amount_deal"))
'                listItem.SubItems(8) = IIf(IsNull(M_Objrs("tenor")), "1", M_Objrs("tenor"))
'                listItem.SubItems(9) = IIf(IsNull(M_Objrs("pembayaran_via")), "", M_Objrs("pembayaran_via"))
'                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tgl_tagih")), "", Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd"))
'                listItem.SubItems(11) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
'                listItem.SubItems(12) = IIf(IsNull(M_Objrs("balance")), "0", M_Objrs("balance"))
'                listItem.SubItems(13) = IIf(IsNull(M_Objrs("pembayaran_awal")), "0", M_Objrs("pembayaran_awal"))
'                listItem.SubItems(14) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
'                listItem.SubItems(15) = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
'                listItem.SubItems(16) = IIf(IsNull(M_Objrs("ndownpay")), "0", M_Objrs("ndownpay"))
'                listItem.SubItems(17) = IIf(IsNull(M_Objrs("ncharge")), "0", M_Objrs("ncharge"))
'                listItem.SubItems(18) = IIf(IsNull(M_Objrs("ndiscountamt")), "0", M_Objrs("ndiscountamt"))
'                listItem.SubItems(19) = IIf(IsNull(M_Objrs("vosbalance")), "", M_Objrs("vosbalance"))
'                listItem.SubItems(20) = IIf(IsNull(M_Objrs("vosprincipal")), "", M_Objrs("vosprincipal"))
'                listItem.SubItems(21) = IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust"))
'                listItem.SubItems(22) = IIf(IsNull(M_Objrs("chkfaxed")), "", M_Objrs("chkfaxed"))
'                listItem.SubItems(23) = IIf(IsNull(M_Objrs("chkwentalking")), "", M_Objrs("chkwentalking"))
'                listItem.SubItems(24) = IIf(IsNull(M_Objrs("chkktp")), "", M_Objrs("chkktp"))
'                listItem.SubItems(25) = IIf(IsNull(M_Objrs("chksup")), "", M_Objrs("chksup"))
'                listItem.SubItems(26) = IIf(IsNull(M_Objrs("chkbillings")), "", M_Objrs("chkbillings"))
'                listItem.SubItems(27) = IIf(IsNull(M_Objrs("chkothers")), "", M_Objrs("chkothers"))
'                listItem.SubItems(28) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
'
'                If IsNull(M_Objrs("dob")) = True Or M_Objrs("dob") = "" Or M_Objrs("dob") = Empty Then
'                    DOB = ""
'                Else
'                    DOB = Format(M_Objrs("dob"), "yyyy-mm-dd")
'                End If
'
'                listItem.SubItems(29) = DOB
'                listItem.SubItems(30) = IIf(IsNull(M_Objrs("ket_other")), "", M_Objrs("ket_other"))
'            M_Objrs.MoveNext
'        Wend
'    End If
'    Set M_Objrs = Nothing

End Sub

Private Sub Check1_Click()
    Dim r As Integer
        
    If Check1.Value = vbChecked Then
        If ListView1.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = True
        Next r
        'Call cmd_count_Click
    Else
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = False
        Next r
        'Call cmd_count_Click
    End If

End Sub

Private Sub CmbJenisRejected_Click()
    Call HeaderLogRejected
    Call IsiLogRejected
End Sub

Private Sub CmbKembalikan_Click()
    Dim cmdsql As String
    Dim K As String
    Dim w As Integer
    
    If LvPTPRejected.ListItems.Count = 0 Then
        MsgBox "Data List PTP Rejected tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    K = MsgBox("Apakah anda yakin akan mengembalikan PTP Rejected ke List PTP Request?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If K = vbNo Then
        MsgBox "Pengembalian data dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvPTPRejected.ListItems.Count
        If LvPTPRejected.ListItems(w).Checked = True Then
            cmdsql = "insert into tblsendptp "
            cmdsql = cmdsql + " select * from tblsendptp_log_reject where id='"
            cmdsql = cmdsql + CStr(LvPTPRejected.ListItems(w).text) + "'"
            M_OBJCONN.execute cmdsql
            
            cmdsql = "delete from tblsendptp_log_reject where id='"
            cmdsql = cmdsql + CStr(LvPTPRejected.ListItems(w).text) + "'"
            M_OBJCONN.execute cmdsql
        End If
    Next w
    
    MsgBox "Data PTP Rejected berhasil dikembalikan ke list request PTP!", vbOKOnly + vbInformation, "Informasi"
    
    Call IsiLogRejected
End Sub

Private Sub CmbTampilkan_Click()
'    If CmbTampilkan.Text = "PTP DISC." Then
'        LvPTP.CheckBoxes = False
'    End If
'    If CmbTampilkan.Text = "PTP NO DISC." Then
'        LvPTP.CheckBoxes = True
'    End If

    If CmbTampilkan.text = "PTP NO DISC." Then
        'CmdApproveByPTP.Visible = False
        cmdapprove.Visible = True
    End If
    If CmbTampilkan.text = "PTP DISC." Then
        If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
            'CmdApproveByPTP.Visible = False
            cmbapprove.Visible = False
        Else
             'CmdApproveByPTP.Visible = True
             'CmdApprove.Visible = False
             cmbapprove.Visible = True
        End If
        
    End If
    If CmbTampilkan.text = "PTP DISC. APPROVED" Then
        'CmdApproveByPTP.Visible = False
    End If
    
    Call HeaderLog
    Call isilog
End Sub

Private Sub cmd_SID_Click()
    Dim K As Integer
    Dim w As String
    Dim r As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    'UPDATED BY RANDY BUAT SID -- REQ BY : JOKO
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            FrmSID.List1.AddItem LvPTP.ListItems(K).ListSubItems(2)
            LvPTP.ListItems(K).Checked = False
        End If
    Next K
  
    FrmSID.Show vbModal
End Sub




Private Sub CmdApprove_Click()
    Dim K As Integer
    Dim w As String
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Or UCase(MDIForm1.Text2.text) = "MANAGER" Then
        If CmbTampilkan.text = "PTP DISC." And _
           LvPTP.SelectedItem.SubItems(4) = "Belum di Approve" Then
            MsgBox "PTP yang akan anda approve adalah PTP Discon! Untuk Meng-approve-nya harus melalui persetujuan SPV, double click data yang akan di approve kemudian Print dan ajukan ke SPV!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
     End If
    
    If CmbTampilkan.text = "PTP DISC." And _
       cmbapprove.text = "" Then
        MsgBox "Approve By, tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    w = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP dan CPA?", vbYesNo + vbQuestion, "Konfirmasi")
    If w = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvPTP.ListItems.Count
    
    'UPDATED BY RANDY BUAT NYATET TGL_APPROVE KARENA SEBELUMNYA KOSONG -- REQ BY : NYOTO
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            cmdsql = "update tblsendptp set tgl_approve=now() where id='"
            cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
            M_OBJCONN.execute cmdsql
        End If
    Next K
    
    cmdapprove.Enabled = False
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            Call BikinCPA(K)
            DoEvents
            Call BikinPTP(K)
            DoEvents
            Call CatetLogApprove(K)
            DoEvents
            Call BikinStatusPTP(K)
            DoEvents
            Call HapusData(K)
            DoEvents
            Call KirimPesan(K)
        End If
    Next K
    Call isilog
    MsgBox "Approve PTP berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
    cmdapprove.Enabled = True
End Sub

Private Sub CmdApproveByPTP_Click()
    Dim cmdsql, w As String
    Dim K As Integer
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Or UCase(MDIForm1.Text2.text) = "MANAGER" Then
        MsgBox "Approve PTP Discon Hanya Boleh dilakukan oleh SPV!", vbOKOnly + vbInformation, "Informasi"
    End If
       
    
    w = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP data yang dicentang?", vbYesNo + vbQuestion, "Konfirmasi")
    If w = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

    PB1.Max = LvPTP.ListItems.Count
    
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            cmdsql = "update tblsendptp set status='1',tgl_approve=now() where id='"
            cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
            M_OBJCONN.execute cmdsql
        End If
    Next K
    Call isilog
    MsgBox "Approve PTP berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdApproveHamanto_Click()
    Dim w As String
    Dim K As Integer
    
    If LvHamanto.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglApprove.ValueIsNull = True Then
        MsgBox "Tanggal Approve tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    w = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP dan CPA?", vbYesNo + vbQuestion, "Konfirmasi")
    If w = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB2.Max = LvHamanto.ListItems.Count
    
    CmdApproveHamanto.Enabled = False
    TxtTglApprove.Enabled = False
    For K = 1 To LvHamanto.ListItems.Count
        PB1.Value = K
        If LvHamanto.ListItems(K).Checked = True Then
            Call BikinCPA_Hamanto(K)
            DoEvents
            Call BikinPTP_Hamanto(K)
            DoEvents
            Call CatetLogApprove_Hamanto(K)
            DoEvents
            Call BikinStatusPTP_Hamanto(K)
            DoEvents
            Call HapusData_Hamanto(K)
            DoEvents
            Call KirimPesan_Hamanto(K)
        End If
    Next K
    Call isilog
    MsgBox "Approve PTP berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
    CmdApproveHamanto.Enabled = True
    TxtTglApprove.Enabled = True
    CmdCariAppHamanto_Click
End Sub

Private Sub CmdApproveVP_Click()
    Dim K As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        If CmbTampilkan.text = "PTP DISC." And _
           LvPTP.SelectedItem.SubItems(4) = "Belum di Approve" Then
            MsgBox "PTP yang akan anda approve adalah PTP Discon! Untuk Meng-approve-nya harus melalui persetujuan SPV, double click data yang akan di approve kemudian Print dan ajukan ke SPV!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
    End If
    
    If CmbTampilkan.text <> "PTP DISC." Then
        MsgBox "Approve By VP hanya untuk PTP Disc.!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    w = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP dan CPA?", vbYesNo + vbQuestion, "Konfirmasi")
    If w = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvPTP.ListItems.Count
    
    CmdApproveVP.Enabled = False
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            Call BikinCPA_AppVP(K)
            DoEvents
            'Call BikinPTP(K)
            'DoEvents
            'Call CatetLogApprove(K)
            'DoEvents
            'Call BikinStatusPTP(K)
            'DoEvents
            'Call HapusData(K)
            'DoEvents
            Call KirimPesan_AppVP(K)
        End If
    Next K
    Call isilog
    MsgBox "Pengajuan CPA  berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
    CmdApproveVP.Enabled = True
End Sub

Private Sub CmdCari_Click()
    Call isilog
End Sub

Private Sub CmdCariAppHamanto_Click()
    Dim cmdsql As String
    Dim M_WHERE As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
  
    
    
    M_WHERE = ""
    
    If TxtCustidHamanto.text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where custid like '%" + CStr(TxtCustidHamanto.text) + "%' "
        Else
            M_WHERE = M_WHERE + " and custid like '%" + CStr(TxtCustidHamanto.text) + "%' "
        End If
    End If
    
    If TxtCariNamaHamanto.text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where vcustname like '%" + CStr(TxtCariNamaHamanto.text) + "%' "
        Else
            M_WHERE = M_WHERE + " and vcustname like '%" + CStr(TxtCariNamaHamanto.text) + "%' "
        End If
    End If
    
    If M_WHERE = "" Then
        M_WHERE = " where sts_app_vp='1' "
    Else
        M_WHERE = M_WHERE + " and sts_app_vp='1' "
    End If
    
    cmdsql = "select * from tblsendptp " + M_WHERE
        
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvHamanto.ListItems.clear
    TxtJmlhAppHamanto.text = M_Objrs.RecordCount
    
    If M_Objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        While Not M_Objrs.EOF
            'On Error Resume Next
            Set listItem = LvHamanto.ListItems.ADD(, , M_Objrs("id"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("vcustname")), "", M_Objrs("vcustname"))
                
                If M_Objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_Objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_Objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                listItem.SubItems(4) = STATUS
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("tgl_approve")), "", Format(M_Objrs("tgl_approve"), "yyyy-mm-dd"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("date_payment_effective")), "", Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd"))
                listItem.SubItems(7) = IIf(IsNull(M_Objrs("total_amount_deal")), "0", M_Objrs("total_amount_deal"))
                listItem.SubItems(8) = IIf(IsNull(M_Objrs("tenor")), "1", M_Objrs("tenor"))
                listItem.SubItems(9) = IIf(IsNull(M_Objrs("pembayaran_via")), "", M_Objrs("pembayaran_via"))
                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tgl_tagih")), "", Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd"))
                listItem.SubItems(11) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
                listItem.SubItems(12) = IIf(IsNull(M_Objrs("balance")), "0", M_Objrs("balance"))
                listItem.SubItems(13) = IIf(IsNull(M_Objrs("pembayaran_awal")), "0", M_Objrs("pembayaran_awal"))
                listItem.SubItems(14) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
                listItem.SubItems(15) = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
                listItem.SubItems(16) = IIf(IsNull(M_Objrs("ndownpay")), "0", M_Objrs("ndownpay"))
                listItem.SubItems(17) = IIf(IsNull(M_Objrs("ncharge")), "0", M_Objrs("ncharge"))
                listItem.SubItems(18) = IIf(IsNull(M_Objrs("ndiscountamt")), "0", M_Objrs("ndiscountamt"))
                listItem.SubItems(19) = IIf(IsNull(M_Objrs("vosbalance")), "", M_Objrs("vosbalance"))
                listItem.SubItems(20) = IIf(IsNull(M_Objrs("vosprincipal")), "", M_Objrs("vosprincipal"))
                listItem.SubItems(21) = IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust"))
                listItem.SubItems(22) = IIf(IsNull(M_Objrs("chkfaxed")), "", M_Objrs("chkfaxed"))
                listItem.SubItems(23) = IIf(IsNull(M_Objrs("chkwentalking")), "", M_Objrs("chkwentalking"))
                listItem.SubItems(24) = IIf(IsNull(M_Objrs("chkktp")), "", M_Objrs("chkktp"))
                listItem.SubItems(25) = IIf(IsNull(M_Objrs("chksup")), "", M_Objrs("chksup"))
                listItem.SubItems(26) = IIf(IsNull(M_Objrs("chkbillings")), "", M_Objrs("chkbillings"))
                listItem.SubItems(27) = IIf(IsNull(M_Objrs("chkothers")), "", M_Objrs("chkothers"))
                listItem.SubItems(28) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                  
                If IsNull(M_Objrs("dob")) = True Or M_Objrs("dob") = "" Or M_Objrs("dob") = Empty Then
                    DOB = ""
                Else
                    DOB = Format(M_Objrs("dob"), "yyyy-mm-dd")
                End If
                 
                listItem.SubItems(29) = DOB
                listItem.SubItems(30) = IIf(IsNull(M_Objrs("ket_other")), "", M_Objrs("ket_other"))
                listItem.SubItems(31) = IIf(IsNull(M_Objrs("payment_handle")), "", M_Objrs("payment_handle"))
                
                listItem.SubItems(32) = IIf(IsNull(M_Objrs("occupation")), "", M_Objrs("occupation"))
                listItem.SubItems(33) = IIf(IsNull(M_Objrs("reason")), "", M_Objrs("reason"))
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing

End Sub

Private Sub CmdCariRejected_Click()
    Call IsiLogRejected
End Sub

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB1.Max = LvPTP.ListItems.Count
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        LvPTP.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdCekAllHamanto_Click()
    Dim K As Integer
    
    If LvHamanto.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB2.Max = LvHamanto.ListItems.Count
    For K = 1 To LvHamanto.ListItems.Count
        PB1.Value = K
        LvHamanto.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdExport_Click()
    'PanelExport.Visible = True
    'Call Export_To_Excel
    Dim xx As Integer
    Dim ceklst As Boolean
    
    ceklst = False
    For xx = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(xx).Checked = True Then
            ceklst = True
            Exit For
        End If
    Next xx
    
    If ceklst Then
        If cmbapprove.text <> "" Then
            Call My_Export_Excel
        Else
            MsgBox "Anda belum memilih data akan di 'Approve By : '", vbCritical + vbOKOnly, "Info"
        End If
    Else
        MsgBox "Anda belum memilih data!", vbOKOnly + vbCritical, "INFO"
    End If
End Sub

Private Sub CmdRefresh_Click()
    Call isilog
End Sub

Private Sub CmdReject_Click()
    Dim K As Integer
    Dim w As String
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    w = MsgBox("Anda yakin akan melakukan Menghapus Request PTP?", vbYesNo + vbQuestion, "Konfirmasi")
    If w = vbNo Then
        MsgBox "Penghapusan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB1.Max = LvPTP.ListItems.Count
    
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            Call CatetLogReject(K)
            Call HapusData(K)
            Call KirimPesanGagal(K)
        End If
    Next K
    
    Call isilog
    MsgBox "Data Berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub cmdsearch_Click()
Dim S As String
Dim sql As String

    If cbsearch.text = "Tanggal Request by Email" Then
        S = "tgl_send_email"
    ElseIf cbsearch.text = "Cust ID" Then
        S = "CustId"
    ElseIf cbsearch.text = "Nama Cust" Then
        S = "vcustname"
    ElseIf cbsearch.text = "Tanggal Approve" Then
        S = "tgl_approve"
    End If
'<9feb2018
sql = "select * from tblsendptp_log_approve where id is not null "
'sql = "select tblsendptp_log_approve.*, a.segment, b.jml from tblsendptp_log_approve left join (select custid,segment from mgm where segment <> '' and segment is not null) a on" & vbCrLf
'sql = sql + " tblsendptp_log_approve.custid = a.custid left join (select * from (select custid, count(custid) as jml from tblsendptp_log_approve group by 1) a) b on tblsendptp_log_approve.custid = b.custid where id is not null "

If cbsearch.text = "Jenis PTP" Or cbsearch.text = "Cust ID" Or cbsearch.text = "Nama Cust" Then
sql = sql + "and " & S & " =  '" & txtsearch.text & "'"
ElseIf cbsearch.text = "Tanggal Approve" Or cbsearch.text = "Tanggal Request by Email" Then
sql = sql + "and " & S & " between '" & Format(date1.Value, "yyyy/mm/dd") & "' and '" & Format(date2.Value, "yyyy/mm/dd") & "' order by tgldata desc limit 300 "
End If

If IsNull(date1.Value) Or IsNull(date2.Value) Then
    MsgBox "Tanggal Wajib Diisi", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
Else
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If

    LvPTPApproved.ListItems.clear
    TxtJmlApproved.text = M_Objrs.RecordCount

    If M_Objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
         While Not M_Objrs.EOF
            'On Error Resume Next
            Set listItem = LvPTPApproved.ListItems.ADD(, , M_Objrs("id"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("vcustname")), "", M_Objrs("vcustname"))
                
                If M_Objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_Objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_Objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
            
         discount = IIf(IsNull(M_Objrs("ndiscountamt")), "0", M_Objrs("ndiscountamt"))
            
            
            If discount > 0 And discount <= 2000000 Then
                S = "Coll SPV"
            ElseIf discount <= 10000000 Then
                S = "Coll Band 6"
            ElseIf discount <= 20000000 Then
                S = "Coll Band 5"
            ElseIf discount <= 30000000 Then
                S = "Coll Band 4"
            ElseIf discount <= 50000000 Then
                S = "Head of Coll"
            ElseIf discount <= 100000000 Then
                S = "Head of CCC"
            ElseIf discount <= 2300000000# Then
                S = "Head of CRM"
            End If
                      
                listItem.SubItems(4) = STATUS
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("tgl_approve")), "", Format(M_Objrs("tgl_approve"), "yyyy-mm-dd"))
                listItem.SubItems(6) = S
                listItem.SubItems(7) = IIf(IsNull(M_Objrs("tgl_send_email")), "", Format(M_Objrs("tgl_send_email"), "yyyy-mm-dd"))
                            
                ' Tandain klo ini belum di send_email 'jejaktian 14032016
                If IIf(IsNull(M_Objrs("tgl_send_email")), 0, M_Objrs("tgl_send_email")) = 0 Then
                    For K = 1 To 7
                        listItem.ListSubItems(K).ForeColor = vbRed
                    Next K
                End If
                ' ------------------------------------------
            
            M_Objrs.MoveNext
        Wend

    End If
    Set M_Objrs = Nothing
    
End Sub

Private Sub cmdsudahemail_Click()
    Dim S As String
    
    For K = 1 To LvPTPApproved.ListItems.Count
        If LvPTPApproved.ListItems(K).Checked = True Then
            S = "update tblsendptp_log_approve set tgl_send_email=now() where id='"
            S = S + CStr(LvPTPApproved.ListItems(K).text) + "'"
            M_OBJCONN.execute S
        End If
    Next K
    
    Call IsiLogApproved
    
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB1.Max = LvPTP.ListItems.Count
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        LvPTP.ListItems(K).Checked = False
    Next K
End Sub



Private Sub CmdUnCekAllHamanto_Click()
    Dim K As Integer
    
    If LvHamanto.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB2.Max = LvHamanto.ListItems.Count
    For K = 1 To LvHamanto.ListItems.Count
        PB1.Value = K
        LvHamanto.ListItems(K).Checked = False
    Next K
End Sub
Private Function Export_To_Excel()
    Dim Strsql          As String
    Dim rs              As ADODB.Recordset
    Dim ExlObj          As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim listcustid      As String
    
    On Error GoTo adderr
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            listcustid = listcustid & ",'" & LvPTP.ListItems(K).SubItems(2) & "'"
        End If
    Next K
    
    listcustid = Mid(listcustid, 2)
    
    Strsql = "select custid,vcustname,'" & cmbapprove.text & "' as Approved FROM tblsendptp WHERE custid in (" & listcustid & ")"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

form_save:
    CD_save.ShowSave
    txtlocation.text = CD_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If txtlocation.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
              CmdExport.Enabled = True
              Exit Function
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo form_save        '-> maka goto form_save
        End If
    End If
 
    If rs.RecordCount > 0 Then
       PB1.Max = rs.RecordCount
    End If

 
    Set ExlObj = CreateObject("Excel.Application")
    Set objBook = ExlObj.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
    On Error GoTo SALAH
    'Proses pengsisian nama field ke excel
    objSheet.Cells(1, 1).Value = "List CPA Approve"
    objSheet.Cells(1, 1).Font.Name = "Verdana"
    objSheet.Cells(1, 1).Font.Bold = True:
    objSheet.Cells(2, 1).Value = "Tanggal : " + Format(Now, "dd-mm-yyyy")
    objSheet.Cells(2, 1).Font.Name = "Verdana"
    objSheet.Cells(2, 1).Font.Bold = True:
    objSheet.Cells(4, 1).Value = "NO"
    objSheet.Cells(4, 2).Value = "CARD NUMBER"
    objSheet.Cells(4, 3).Value = "CH NAME"
    objSheet.Cells(4, 4).Value = "APPROVED"
    objSheet.Cells(4, 5).Value = "ADMIN CREATED"
    objSheet.Cells(4, 6).Value = "RECEIVED BY" 'Dikosongkan
    objSheet.Cells(4, 7).Value = "BAN 6"
    objSheet.Cells(4, 8).Value = "BAN 5"
    objSheet.Cells(4, 9).Value = "BAN 4"
    objSheet.Cells(4, 10).Value = "BAN 3"
    objSheet.Cells(4, 11).Value = "BAN 2"
    objSheet.Cells(4, 12).Value = "BAN 1"
    objSheet.Cells(4, 13).Value = "ADMIN RECEIVED"
    objSheet.Cells(4, 14).Value = "SENT BY"
'    objSheet.Cells(4, 4).Value = "DOB":      objSheet.Cells(4, 5).Value = "STATUS PTP"
'    objSheet.Cells(4, 6).Value = "TOTAL PAYMENT":      objSheet.Cells(4, 7).Value = "DOWN PAYMENT"
'    objSheet.Cells(4, 8).Value = "LPD FROM PAYMENT":      objSheet.Cells(4, 9).Value = "LPA FROM PAYMENT"
                
'select dpropsal,vcustid,vcustname,dob,status_ptp,nttlpayment,ndownpay,lpd_from_payment,lpa_from_payment
    
    objSheet.Range("A5").CopyFromRecordset rs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs txtlocation.text, xlWorkbookNormal
    ExlObj.Quit
    Set ExlObj = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    

    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    PB1.Value = 0
    Command1.Enabled = True
    
    
    
    Set rs = Nothing

    StartMeUp (txtlocation.text)

    txtlocation.text = ""
SALAH:
    Exit Function
    
        
adderr:
    If err.number = -2147217900 Then
    On Error Resume Next
    Resume
    End If
    MsgBox err.Description



End Function

Private Sub cmrefresh_Click()
Unload Me
FrmListRequestPTP.Show vbModal
End Sub

Private Sub Command1_Click()
    Export_To_Excel
End Sub

Private Sub Command2_Click()
PanelExport.Visible = False
End Sub

Private Sub Command3_Click()
    headerlistappptp
    isilistappptp
End Sub

Private Sub Command4_Click()
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data Is Empty!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            cek = cek + 1
        End If
    Next i
        
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    Call funcapp
    MsgBox "Approved"
    isilistappptp
End Sub

Private Sub funcapp()
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            CustId = cnull(ListView1.ListItems(i).text)
            ID = cnull(ListView1.ListItems(i).SubItems(1))
            promdate = cnull(ListView1.ListItems(i).SubItems(3))
            prompay = cnull(ListView1.ListItems(i).SubItems(4))
            stsdata = cnull(ListView1.ListItems(i).SubItems(5))
            via = cnull(ListView1.ListItems(i).SubItems(6))
            agent = cnull(ListView1.ListItems(i).SubItems(7))
            TEAM = cnull(ListView1.ListItems(i).SubItems(8))
            req = cnull(ListView1.ListItems(i).SubItems(9))
            lpd = cnull(ListView1.ListItems(i).SubItems(10))
            lpa = cnull(ListView1.ListItems(i).SubItems(11))
            tagih = cnull(ListView1.ListItems(i).SubItems(12))
            Tenor = cnull(ListView1.ListItems(i).SubItems(13))
            
            cmdsql = "select * from tblnegoptp where custid='"
            cmdsql = cmdsql + CStr(CustId) + "' order by promisedate desc limit 1"
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs.RecordCount > 0 Then
                zzz = M_Objrs("id")
            End If
        
            running = "update tblnegoptp set promisedate='" + promdate + "',promisepay = '" + prompay + "' where id = " & zzz & ";"
            
            cmdsql = "select * from tblcpa where vcustid='"
            cmdsql = cmdsql + CustId + "' order by nid desc limit 1"
            Set M_Objrs_Cek_CPA = New ADODB.Recordset
            M_Objrs_Cek_CPA.CursorLocation = adUseClient
            M_Objrs_Cek_CPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_Cek_CPA.RecordCount > 0 Then
                If Val(M_Objrs_Cek_CPA("nttlpayment")) > 0 Then
                    AmountDeal = M_Objrs_Cek_CPA("nttlpayment")
                End If
            End If
            
            Set M_Objrs_Cek_CPA = Nothing
                 
           If stsdata = "PTP-NEW" Then
                Cmdsql_Cek_status = "select * from mgm where custid='"
                Cmdsql_Cek_status = Cmdsql_Cek_status + CustId + "'"
                Set M_Objrs_Cek_Status = New ADODB.Recordset
                M_Objrs_Cek_Status.CursorLocation = adUseClient
                M_Objrs_Cek_Status.Open Cmdsql_Cek_status, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                If M_Objrs_Cek_Status.RecordCount > 0 Then
                    If M_Objrs_Cek_Status("tglptpnew") = "" Or IsNull(M_Objrs_Cek_Status("tglptpnew")) = True _
                       Or M_Objrs_Cek_Status("tglptpnew") = Empty Then
                        TglPTPNew = "now()"
                     Else
                        TglPTPNew = "'" + CStr(Format(M_Objrs_Cek_Status("tglptpnew"), "yyyy-mm-dd")) + "'"
                     End If
                End If
                
                Set M_Objrs_Cek_Status = Nothing
            
                running = running + "update mgm set dateptpnew='"
                running = running + promdate + "',tgl_tagih='"
                running = running + tagih + "', "
                running = running + " tglallptp='"
                running = running + promdate + "',f_cek_new='PTP-NE',"
                running = running + "kethslkerja_new='PTP-NEW',kethslkerjadesc_new='PTP-NEW',ptpvia='"
                running = running + via + "',ptpdesc='PTP-NEW', dateptp='"
                running = running + promdate + "',tglptpnew=" + TglPTPNew
                running = running + ",tenor='"
                running = running + Tenor + "'" 'ttlptp='"
                'running = running + CStr(AmountDeal) + "' "
                running = running + "where custid='"
                running = running + CustId + "';" & vbCrLf
                DoEvents
            End If
            
            If stsdata = "PTP-POP" Then
                running = running + "update mgm set dateptp='"
                running = running + promdate + "',tgl_tagih='"
                running = running + tagih + "',tglallptp='"
                running = running + promdate + "',f_cek_new='PTP-PO',"
                
                running = running + "kethslkerja_new='PTP-POP',kethslkerjadesc_new='PTP-POP',ptpvia='"
                running = running + via + "',ptpdesc='PTP-POP',"
                running = running + "tenor='"
                running = running + Tenor + "'" ',ttlptp='"
                'running = running + CStr(AmountDeal) + "'"
                running = running + " where custid='"
                running = running + CustId + "';" & vbCrLf
            End If
        
            If via = "" Then
                running = running + "update mgm set ptpvia='ATM LAINNYA' where custid='"
                running = running + CustId + "';"
            End If
            
            'untuklog
            running = running + "insert into tblnegoptp_temp_app_log select * from tblnegoptp_temp_app where id = '" + ID + "';" & vbCrLf
            running = running + "update tblnegoptp_temp_app_log set appby = '" + MDIForm1.Text1.text + "', appdate = now(), status = 'Approve' where id = '" + ID + "';" & vbCrLf
            running = running + "delete from tblnegoptp_temp_app where id = '" + ID + "';"
            running = running
            M_OBJCONN.execute running
        
            'message dan hst
            StatusRemarks = "Change PTP Approve by: " & MDIForm1.Text1.text & ""
        
        cmdsql = "insert into mgm_hst(custid,agent,hst,f_cek_new,user_log) values ('"
        cmdsql = cmdsql + CStr(CustId) + "','"
        cmdsql = cmdsql + CStr(agent) + "','"
        cmdsql = cmdsql + StatusRemarks + "','"
        cmdsql = cmdsql + stsdata + "','"
        cmdsql = cmdsql + CStr(MDIForm1.Text1.text) + "')"
        M_OBJCONN.execute cmdsql
        
        Remarks = "Change PTP untuk custid: " & CustId & " telah di approve!"
    
        cmdsql = "insert into msgtbl "
        cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
        cmdsql = cmdsql + agent + "','"
        cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        cmdsql = cmdsql + Remarks + "')"
        M_OBJCONN.execute cmdsql
            
        cmdsql = "select team from usertbl where userid='"
        cmdsql = cmdsql + CStr(agent) + "' "
        cmdsql = cmdsql + " and team is not null "
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "(recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + CStr(Trim(M_Objrs("team"))) + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.Text1.text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + Remarks + "')"
            M_OBJCONN.execute cmdsql
        End If
        
        Set M_Objrs = Nothing
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        
        End If
    Next i
            
End Sub

Private Sub Command5_Click()

    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data Is Empty!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If

    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            cek = cek + 1
        End If
    Next i

    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            CustId = cnull(ListView1.ListItems(i).text)
            ID = cnull(ListView1.ListItems(i).SubItems(1))
            promdate = cnull(ListView1.ListItems(i).SubItems(3))
            prompay = cnull(ListView1.ListItems(i).SubItems(4))
            stsdata = cnull(ListView1.ListItems(i).SubItems(5))
            via = cnull(ListView1.ListItems(i).SubItems(6))
            agent = cnull(ListView1.ListItems(i).SubItems(7))
            TEAM = cnull(ListView1.ListItems(i).SubItems(8))
            req = cnull(ListView1.ListItems(i).SubItems(9))
            lpd = cnull(ListView1.ListItems(i).SubItems(10))
            lpa = cnull(ListView1.ListItems(i).SubItems(11))
            tagih = cnull(ListView1.ListItems(i).SubItems(12))
            Tenor = cnull(ListView1.ListItems(i).SubItems(13))
            
            running = "insert into tblnegoptp_temp_app_log select * from tblnegoptp_temp_app where id = " & ID & ";" & vbCrLf
            running = running + "update tblnegoptp_temp_app_log set appby = '" + MDIForm1.Text1.text + "', appdate = now(), status = 'Reject' where id = " & ID & ";" & vbCrLf
            running = running + "delete from tblnegoptp_temp_app where id = " & ID & ";"
            
            M_OBJCONN.execute running
            
            'Message dan hst
            
            StatusRemarks = "Change PTP Rejected by: " & MDIForm1.Text1.text & ""
        
            cmdsql = "insert into mgm_hst(custid,agent,hst,f_cek_new,user_log) values ('"
            cmdsql = cmdsql + CStr(CustId) + "','"
            cmdsql = cmdsql + CStr(agent) + "','"
            cmdsql = cmdsql + StatusRemarks + "','"
            cmdsql = cmdsql + stsdata + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Text1.text) + "')"
            M_OBJCONN.execute cmdsql
            
            Remarks = "Change PTP untuk custid: " & CustId & " telah di Reject!"
        
            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + agent + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.Text1.text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + Remarks + "')"
            M_OBJCONN.execute cmdsql
                
            cmdsql = "select team from usertbl where userid='"
            cmdsql = cmdsql + CStr(agent) + "' "
            cmdsql = cmdsql + " and team is not null "
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs.RecordCount > 0 Then
                cmdsql = "insert into msgtbl "
                cmdsql = cmdsql + "(recipient, datetime, sender, sentfrom, msg) values ('"
                cmdsql = cmdsql + CStr(Trim(M_Objrs("team"))) + "','"
                cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
                cmdsql = cmdsql + MDIForm1.Text1.text + "','"
                cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
                cmdsql = cmdsql + Remarks + "')"
                M_OBJCONN.execute cmdsql
            End If
            
            Set M_Objrs = Nothing
            
         '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
         End If
    Next i
            MsgBox "Rejected"
            isilistappptp
End Sub

Private Sub Command6_Click()
    headerlistappptp_log
    isilisthstchange
End Sub

Private Sub Command7_Click()
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
                    objExcelSheet.Cells(Row, col).Value = "'" + ListView1.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = ListView1.ListItems(Row - 1).SubItems(col - 1)
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

Private Sub Command8_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    'On Error GoTo zzz
    If LvPTPApproved.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To LvPTPApproved.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = LvPTPApproved.ColumnHeaders(col)
        Next
     
        For Row = 2 To LvPTPApproved.ListItems.Count + 1
            For col = 1 To LvPTPApproved.ColumnHeaders.Count
            If col = 1 Then
                If col >= 7 And col <= 14 Then
                    objExcelSheet.Cells(Row, col).Value = LvPTPApproved.ListItems(Row - 1).text
                Else
                    objExcelSheet.Cells(Row, col).Value = "'" + LvPTPApproved.ListItems(Row - 1).text
                End If
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = LvPTPApproved.ListItems(Row - 1).SubItems(col - 1)
                    hasil1 = Replace(hasil1, "=", "")
                    'On Error GoTo bawah
                    If col >= 7 And col <= 14 Then
                        objExcelSheet.Cells(Row, col).Value = hasil1
                    Else
                        objExcelSheet.Cells(Row, col).Value = "'" + hasil1
                    End If
'bawah:
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

Private Sub Form_Load()
    CmbTampilkan.text = "PTP NO DISC."
    CmbJenisRejected.text = "PTP NO DISC."
    
    Call HeaderLog
    Call isilog
    
    Call HeaderLogRejected
    Call IsiLogRejected
    
    Call HeaderLogApproved
    Call IsiLogApproved
    
    '@@221012 Bikin header log pak hamanto
    Call HeaderAppHamanto
        
    PanelExport.Visible = False
    'To Be Approved By Pak Hamanto
    
    Call headerlistappptp
    
    If MDIForm1.Text2.text <> "Supervisor" And MDIForm1.Text2.text <> "Manager" Then
        SSTab1.TabVisible(4) = False
    End If
End Sub

Private Sub BikinCPA(K As Integer)
    Dim cmdsql As String
    Dim Remarks As String
    Dim M_Objrs_Cek_Type As ADODB.Recordset
    Dim TypeAcc As String
    
    TypeAcc = ""
    
    '@@13022013 Cek type account dulu nih .. pil/card
    cmdsql = "select acc_type from mgm where custid='"
    cmdsql = cmdsql & CStr(LvPTP.ListItems(K).SubItems(2)) & "'"
    Set M_Objrs_Cek_Type = New ADODB.Recordset
    M_Objrs_Cek_Type.CursorLocation = adUseClient
    M_Objrs_Cek_Type.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek_Type.RecordCount > 0 Then
        TypeAcc = IIf(IsNull(M_Objrs_Cek_Type("acc_type")), "", M_Objrs_Cek_Type("acc_type"))
    End If
    
    Set M_Objrs_Cek_Type = Nothing
    
    Call Cari_LPD_LPA_Payment(K)
    
    cmdsql = "insert into tblcpa (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
    cmdsql = cmdsql + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
    cmdsql = cmdsql + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
    cmdsql = cmdsql + "chksup,chkbillings,chkothers,lpd_from_payment,lpa_from_payment,"
    cmdsql = cmdsql + "f_system,dob,status_ptp,ketother "
    
    '@@19062012 Jika Status PTP DISCON Catat Approvenya
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = cmdsql + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    'Catet Juga yang PTP No Discon 20062012
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = cmdsql + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    '@@16-07-2012 Buat Catet Payment Handle
    cmdsql = cmdsql + " ,vpaymenthandle,voccupation,vreason "
    
    cmdsql = cmdsql + ") values ("
    cmdsql = cmdsql + "now(),'"
    'Cmdsql = Cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "','CARD','"
    cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
    cmdsql = cmdsql + TypeAcc + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(15)), "0", Replace(LvPTP.ListItems(K).SubItems(15), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(16)), "0", Replace(LvPTP.ListItems(K).SubItems(16), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(17)), "0", Replace(LvPTP.ListItems(K).SubItems(17), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(18)), "0", Replace(LvPTP.ListItems(K).SubItems(18), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "',"
    cmdsql = cmdsql + "now(),'"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(3)), "", LvPTP.ListItems(K).SubItems(3))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(21)), "", LvPTP.ListItems(K).SubItems(21))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "0", Replace(LvPTP.ListItems(K).SubItems(12), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(11)), "0", Replace(LvPTP.ListItems(K).SubItems(11), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(22)), "", LvPTP.ListItems(K).SubItems(22))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(23)), "", LvPTP.ListItems(K).SubItems(23))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(24)), "", LvPTP.ListItems(K).SubItems(24))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(25)), "", LvPTP.ListItems(K).SubItems(25))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(26)), "", LvPTP.ListItems(K).SubItems(26))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(27)), "", LvPTP.ListItems(K).SubItems(27))) + "',"
    cmdsql = cmdsql + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
    cmdsql = cmdsql + CStr(TxtLPAPayment.Value) + "','1',"
    '@@20062012 Tambahkan DOB dan Status PTP
    cmdsql = cmdsql + IIf(LvPTP.ListItems(K).SubItems(29) = "", "null", "'" + LvPTP.ListItems(K).SubItems(29) + "'")
    cmdsql = cmdsql + ",'" + LvPTP.ListItems(K).SubItems(1) + "',' "
    '@@21062012 Tambahkan Keterangan Other
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(30)), "", LvPTP.ListItems(K).SubItems(30)) + "' "
    
    '@@19062012 Buat nyatet approvenya
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = cmdsql + ",now(),'1','"
        cmdsql = cmdsql + Trim(cmbapprove.text) + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "'"
     End If
     
     'Buat nyatet yang jenisnya PTP NO Discount.
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = cmdsql + ",now(),'1','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "'"
     End If
    
    cmdsql = cmdsql + ",'"
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(31)), "", LvPTP.ListItems(K).SubItems(31)) + "','"
    
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(32)), "", LvPTP.ListItems(K).SubItems(32)) + "','"
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(33)), "", LvPTP.ListItems(K).SubItems(33)) + "')"
    DoEvents
    M_OBJCONN.execute cmdsql
    
    '@@19062012 Bikin Remarks untuk CPA
     '@@11092012 Tulis Remarks baik untuk yang ptp discon/no discon
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        Remarks = "PtpDisc-"
     Else
        Remarks = "PTPNoDisc-"
     End If
        Remarks = Remarks + "App By:" + cmbapprove.text + "-"
        Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(7)), "", LvPTP.ListItems(K).SubItems(7))) + " -"
        Remarks = Remarks + "Instl: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + " -"
        Remarks = Remarks + "From Bal.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "", LvPTP.ListItems(K).SubItems(12))) + " -"
        Remarks = Remarks + "From Prin.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(14)), "", LvPTP.ListItems(K).SubItems(14))) + " -"
        Remarks = Remarks + "%Balance: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "% -"
        Remarks = Remarks + "%Principal: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "% #USER LOG:" + MDIForm1.Text1.text
        
        cmdsql = "insert into mgm_hst (custid, agent, products, "
        cmdsql = cmdsql + "hst,user_log) values ('"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(28) + "','"
        cmdsql = cmdsql + "Collection" + "','"
        cmdsql = cmdsql + Remarks + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "')"
        M_OBJCONN.execute cmdsql
    
    
    '@@25072012,Update yang approve dan tanggal proposalnya di tabel tblsendptp jika PTP discount
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='"
        cmdsql = cmdsql + CStr(Trim(cmbapprove.text)) + "', log_approve='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "' where id='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.execute cmdsql
    End If
    
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "', log_approve='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "' where id='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.execute cmdsql
    End If
End Sub

'@@ 16-03-2011, Ini buat nyari LPD dan LPA terakhir dari tabel lunas
Private Sub Cari_LPD_LPA_Payment(K As Integer)
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    StatusPTP = ""
    TxtLPDPayment.text = ""
    TxtLPAPayment.Value = "0"
    
'    cmdsql = "select paydate,payment from tbllunas where custid='"
'    cmdsql = cmdsql + Trim(LvPTP.ListItems(K).SubItems(2)) + "' order by paydate desc limit 1 "
    
    cmdsql = " select paydate,payment,case when to_char(paydate,'yyyy-mm') = to_char(now() - interval '1 month','yyyy-mm') then 1 else 0 end zz from tbllunas  where custid  = '" & Trim(LvPTP.ListItems(K).SubItems(2)) & "' order by paydate desc limit 1"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        If M_Objrs!zz = 1 Then
            GoTo bawah
        End If
    End If

    cmdsql = " select paydate,payment,case when to_char(paydate,'yyyy-mm') = to_char(now() - interval '2 month','yyyy-mm') then 1 else 0 end zz from tbllunas  where custid  = '" & Trim(LvPTP.ListItems(K).SubItems(2)) & "' order by paydate desc limit 1"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        If M_Objrs!zz = 1 Then
            GoTo bawah
        End If
    End If
    
    cmdsql = " select paydate,payment,case when to_char(paydate,'yyyy-mm') = to_char(now() - interval '3 month','yyyy-mm') then 1 else 0 end zz from tbllunas  where custid  = '" & Trim(LvPTP.ListItems(K).SubItems(2)) & "' order by paydate desc limit 1"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    GoTo bawahs
    
        If M_Objrs.RecordCount > 0 Then
bawah:
            StatusPTP = "PTP-POP"
            TxtLPDPayment.text = IIf(IsNull(M_Objrs("paydate")), "", Format(M_Objrs("paydate"), "yyyy-mm-dd"))
            TxtLPAPayment.Value = IIf(IsNull(M_Objrs("payment")), "0", M_Objrs("payment"))
            LpdPayment = "'" + TxtLPDPayment.text + "'"
        Else
bawahs:
            StatusPTP = "PTP-NEW"
            'LpdPayment = "null"
            TxtLPDPayment.text = ""
            TxtLPAPayment.Value = "0"
        End If
    Set M_Objrs = Nothing
End Sub

Private Sub BikinPTP(K As Integer)
    Dim cmdsql As String
    Dim i As Integer
    Dim M_Objrs_Cek_Tgl As ADODB.Recordset
    Dim jumlah_tenor As Integer
    
    'Tambahan Randy 7April2015 Untuk ambil jumlah tenor sebagai validasi
    jumlah_tenor = Val(LvPTP.ListItems(K).SubItems(8))
    
    bcekptp = True
    
        'Jika Tenor=1
        If jumlah_tenor = 1 Then
                  
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp where custid='"
                cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(6)) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                  
            jatuhtempo = LvPTP.ListItems(K).SubItems(6)
            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.execute cmdsql
            
            
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
'                '@@14-04-2012 Cek Data
'                CMDSQL = "select * from tblnegoptp_log where custid='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(6)) + "'"
'                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
'                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
'                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
'                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
'                    While Not M_Objrs_Cek_Tgl.EOF
'                        CMDSQL = "delete from tblnegoptp_log where id='"
'                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
'                        M_OBJCONN.Execute CMDSQL
'                        M_Objrs_Cek_Tgl.MoveNext
'                    Wend
'                End If
'                Set M_Objrs_Cek_Tgl = Nothing
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
            
            
            ' isi ke tbl log_ptp
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'" + CStr(LvPTP.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.execute cmdsql
                
        Else
            'Untuk Tenor yang lebih dari 1
                        
                'Hapus Reserved Data
                cmdsql = "delete from tblreserve where custid='"
                cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
                M_OBJCONN.execute cmdsql
                        
                jatuhtempo = CStr(LvPTP.ListItems(K).SubItems(6))
            
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp where custid='"
                cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.execute cmdsql
            
            
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
'                '@@14-04-2012 Cek Data
'                CMDSQL = "select * from tblnegoptp_log where custid='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(6)) + "'"
'                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
'                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
'                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
'                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
'                    While Not M_Objrs_Cek_Tgl.EOF
'                        CMDSQL = "delete from tblnegoptp_log where id='"
'                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
'                        M_OBJCONN.Execute CMDSQL
'                        M_Objrs_Cek_Tgl.MoveNext
'                    Wend
'                End If
'                Set M_Objrs_Cek_Tgl = Nothing
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
            
            
            'isi ke tbl log_ptp
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'" + CStr(LvPTP.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.execute cmdsql
                
                
            n = 0
            
            Call HitungInstallmentPtp(K)
            
            For i = 1 To (Val(LvPTP.ListItems(K).SubItems(8)))
                n = n + 1
                'JMLPAY = ((.TxtPayment - txtPembayaranAwal.Value) - PaymentTenor) / (.txttenor.Value - 1)
                JmlPay = PaymentTenor
                Vrdate = DateAdd("m", n, Format(LvPTP.ListItems(K).SubItems(6), "yyyy-mm-dd"))
                    
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblreserve where custid='"
                cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblreserve where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                    
                    cmdsql = "INSERT INTO tblreserve "
                    cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                    cmdsql = cmdsql + "VALUES "
                    cmdsql = cmdsql + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
                    cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
                    cmdsql = cmdsql + "now(), "
                    cmdsql = cmdsql + "'IPO')"
                    M_OBJCONN.execute cmdsql
  
                    cmdsql = "INSERT INTO TblNegoptp_log "
                    cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
                    cmdsql = cmdsql + "VALUES "
                    cmdsql = cmdsql + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
                    cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
                    cmdsql = cmdsql + "now(), "
                    cmdsql = cmdsql + "'" + CStr(LvPTP.ListItems(K).SubItems(28)) + "','R')"
                    M_OBJCONN.execute cmdsql

            'INSERT KE TABEL PTP-REGULER(Randy07-04-2015)
            cmdsql = "select * from tblnegoptp_reguler where custid='"
            cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
            cmdsql = cmdsql + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
            Set M_Objrs_Cek_Tgl = New ADODB.Recordset
            M_Objrs_Cek_Tgl.CursorLocation = adUseClient
            M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                While Not M_Objrs_Cek_Tgl.EOF
                    cmdsql = "delete from tblnegoptp_reguler where id='"
                    cmdsql = cmdsql + CStr(cnull(M_Objrs_Cek_Tgl("id"))) + "'"
                    M_OBJCONN.execute cmdsql
                    M_Objrs_Cek_Tgl.MoveNext
                Wend
            End If
            Set M_Objrs_Cek_Tgl = Nothing
        
            cmdsql = "INSERT INTO tblnegoptp_reguler"
            cmdsql = cmdsql + "(custid, balance, PromiseDate, Promisepay, inputdate, type, tenor, down_payment, agent, keterangan_ptp) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + " '" + CStr(LvPTP.ListItems(K).SubItems(7)) + "',"
            cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'Reguler',"
            cmdsql = cmdsql + " '" + CStr(LvPTP.ListItems(K).SubItems(8)) + "',"
            cmdsql = cmdsql + " '" + CStr(LvPTP.ListItems(K).SubItems(16)) + "',"
            cmdsql = cmdsql + " '" + CStr(LvPTP.ListItems(K).SubItems(28)) + "', "
            cmdsql = cmdsql + "'PTP-NEW')"
            M_OBJCONN.execute cmdsql
       Next i
       End If
    
    
    PaymentTenor = 0
    
    'MsgBox "PTP berhasil ditambahkan!", vbOKOnly + vbInformation, "Informasi"
End Sub


'@@22-09-2011 Hitung InstallmentPtp
Private Sub HitungInstallmentPtp(K As Integer)
    Dim installment As Double
    
        If Val(LvPTP.ListItems(K).SubItems(8)) = 0 Or Val(LvPTP.ListItems(K).SubItems(8)) = 1 Then
            installment = Val(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) / 1
        Else
            installment = (Val(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) - Val(Replace(LvPTP.ListItems(K).SubItems(13), ",", ""))) / (Val(LvPTP.ListItems(K).SubItems(8)) - 1)
        End If
        PaymentTenor = Ceiling(installment)
End Sub

Private Sub CatetLogApprove(K As Integer)
    Dim cmdsql As String
    cmdsql = "update tblsendptp set status = 1 where id = '" + CStr(LvPTP.ListItems(K).text) + "'"
    DoEvents
    M_OBJCONN.execute cmdsql
    cmdsql = "insert into tblsendptp_log_approve "
    cmdsql = cmdsql + "select * from tblsendptp where id='"
    cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
    DoEvents
    M_OBJCONN.execute cmdsql
End Sub

Private Sub CatetLogReject(K As Integer)
    Dim cmdsql As String
    
    '@@25072012 Catet nih siapa yang melakukan reject
    cmdsql = "update tblsendptp set tgl_proposal=now(),log_approve='"
    cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "' where id='"
    cmdsql = cmdsql + CStr(Trim(LvPTP.ListItems(K).text)) + "'"
    M_OBJCONN.execute cmdsql
    
    cmdsql = "insert into tblsendptp_log_reject "
    cmdsql = cmdsql + "select * from tblsendptp where id='"
    cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
    M_OBJCONN.execute cmdsql
End Sub

Private Sub BikinStatusPTP(K As Integer)
    Dim cmdsql As String
    Dim Cmdsql_Cek As String
    Dim StatusRemarks As String
    Dim M_Objrs_Cek As ADODB.Recordset
    Dim AmountNew As Double
    
    AmountNew = 0
    
    Cmdsql_Cek = "select * from tblnegoptp where custid='"
    Cmdsql_Cek = Cmdsql_Cek + CStr(LvPTP.ListItems(K).SubItems(2)) + "' order by id desc limit 1"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Cek.RecordCount > 0 Then
        AmountNew = Val(IIf(IsNull(M_Objrs_Cek("promisepay")), "0", M_Objrs_Cek("promisepay")))
    Else
       AmountNew = 0
    End If
    
    'Jika StatusPTP=PTP NEW
    If StatusPTP = "PTP-NEW" Then
        Dim M_Objrs_Cek_Status As ADODB.Recordset
        Dim Cmdsql_Cek_status As String
        Dim TglPTPNew As String
        
        'Cari apakah sebelumnya status data=ptp new, jika iya maka tglptpnew tidak usah diupdate
        'Tapi jika status sebelumnya bukan ptp new maka update tglptpnew=now
        Cmdsql_Cek_status = "select * from mgm where custid='"
        Cmdsql_Cek_status = Cmdsql_Cek_status + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
        Set M_Objrs_Cek_Status = New ADODB.Recordset
        M_Objrs_Cek_Status.CursorLocation = adUseClient
        M_Objrs_Cek_Status.Open Cmdsql_Cek_status, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_Status.RecordCount > 0 Then
            If M_Objrs_Cek_Status("tglptpnew") = "" Or IsNull(M_Objrs_Cek_Status("tglptpnew")) = True _
               Or M_Objrs_Cek_Status("tglptpnew") = Empty Then
                TglPTPNew = "now()"
             Else
                TglPTPNew = "'" + CStr(Format(M_Objrs_Cek_Status("tglptpnew"), "yyyy-mm-dd")) + "'"
             End If
        End If
        
        Set M_Objrs_Cek_Status = Nothing
    
        cmdsql = "update mgm set dateptpnew='"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(6) + "',tgl_tagih='"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(10) + "', amountnew='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',tglallptp='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + "',tglallptp='"
        
        '@@20062012, amountnew ambil dari negoptp terakhir aja deh....
        cmdsql = cmdsql + CStr(AmountNew) + "',tglallptp='"
        
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(6) + "',f_cek_new='PTP-NE',"
        cmdsql = cmdsql + "tglincoming=now(),ttlptp='"
        cmdsql = cmdsql + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',"
        cmdsql = cmdsql + "kethslkerja_new='PTP-NEW',kethslkerjadesc_new='PTP-NEW',ptpvia='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-NEW', dateptp='"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(6) + "',tglptpnew=" + TglPTPNew
        cmdsql = cmdsql + ",tenor='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(8)) + "' "
        cmdsql = cmdsql + "where custid='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
        DoEvents
        M_OBJCONN.execute cmdsql
        
    End If
    
    If StatusPTP = "PTP-POP" Then
        cmdsql = "update mgm set dateptp='"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(6) + "',tgl_tagih='"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(10) + "',tglallptp='"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(6) + "',f_cek_new='PTP-PO',"
        cmdsql = cmdsql + "tglincoming=now(),ttlptp='"
        cmdsql = cmdsql + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',"
        cmdsql = cmdsql + "kethslkerja_new='PTP-POP',kethslkerjadesc_new='PTP-POP',ptpvia='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-POP',amountptp='"
        cmdsql = cmdsql + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',tenor='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(8)) + "' "
        cmdsql = cmdsql + "where custid='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
        M_OBJCONN.execute cmdsql
    End If
    
     '@@19062012 Bikin Remark Status PTP
        StatusRemarks = "PTP Approve by: " & MDIForm1.Text1.text & "/"
        StatusRemarks = StatusRemarks & "Jenis PTP:" & StatusPTP & "/"
        StatusRemarks = StatusRemarks & "Amount PTP:"
        StatusRemarks = StatusRemarks & CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "PTP Via:" & ""
        StatusRemarks = StatusRemarks & CStr(Replace(LvPTP.ListItems(K).SubItems(9), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "Date PTP:" & Format(LvPTP.ListItems(K).SubItems(6), "yyyy-mm-dd")
        
        cmdsql = "insert into mgm_hst(custid,agent,hst,f_cek_new,user_log) values ('"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(28)) + "','"
        cmdsql = cmdsql + StatusRemarks + "','"
        cmdsql = cmdsql + StatusPTP + "','"
        cmdsql = cmdsql + CStr(MDIForm1.Text1.text) + "')"
        M_OBJCONN.execute cmdsql
        
   Set M_Objrs_Cek = Nothing
        
End Sub

Private Sub HapusData(K As Integer)
    Dim cmdsql As String
    
    cmdsql = "delete from tblsendptp where id='"
    cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
    M_OBJCONN.execute cmdsql
End Sub

Private Sub KirimPesan(K As Integer)
    Dim cmdsql As String
    Dim Remarks As String
    Dim M_Objrs As ADODB.Recordset
    
    Remarks = "Pembuatan PTP untuk custid: " & LvPTP.ListItems(K).SubItems(2) & " telah di approve!"
    
    cmdsql = "insert into msgtbl "
    cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
    cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(28) + "','"
    cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
    cmdsql = cmdsql + MDIForm1.Text1.text + "','"
    cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    cmdsql = cmdsql + Remarks + "')"
    M_OBJCONN.execute cmdsql
        
        
    '@@19072012 Kirim Pesan Buat Ke TL
    'Cari Nama TLNYA
    cmdsql = "select team from usertbl where userid='"
    cmdsql = cmdsql + CStr(Trim(LvPTP.ListItems(K).SubItems(28))) + "' "
    cmdsql = cmdsql + " and team is not null "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        cmdsql = "insert into msgtbl "
        cmdsql = cmdsql + "(recipient, datetime, sender, sentfrom, msg) values ('"
        cmdsql = cmdsql + CStr(Trim(M_Objrs("team"))) + "','"
        cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        cmdsql = cmdsql + Remarks + "')"
        M_OBJCONN.execute cmdsql
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub KirimPesanGagal(K As Integer)
    Dim cmdsql As String
    Dim Remarks As String
    
    Remarks = "Pembuatan PTP untuk custid: " & LvPTP.ListItems(K).SubItems(2) & " telah di reject!"
    
    cmdsql = "insert into msgtbl "
    cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
    cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(28) + "','"
    cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
    cmdsql = cmdsql + MDIForm1.Text1.text + "','"
    cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    cmdsql = cmdsql + Remarks + "')"
    
    M_OBJCONN.execute cmdsql
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    IndexColumnHEader = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = ListView1.SelectedItem.text
        Me.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If

End Sub

Private Sub LvPTP_DblClick()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim TypeAcc As String
    
    TypeAcc = ""
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    cmdsql = "select * from mgm where custid='"
    cmdsql = cmdsql + CStr(LvPTP.SelectedItem.SubItems(2)) + "'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        FrmViewPTP.dtcardopen.Value = IIf(IsNull(M_Objrs("opendate")), "", Format(M_Objrs("opendate"), "dd/mm/yyyy"))
        FrmViewPTP.dwo.Value = IIf(IsNull(M_Objrs("b_d")), "", Format(M_Objrs("b_d"), "dd/mm/yyyy"))
        FrmViewPTP.txtregion.text = IIf(IsNull(M_Objrs("region")), "", M_Objrs("region"))
    End If
    
    TypeAcc = IIf(IsNull(M_Objrs("acc_type")), "", M_Objrs("acc_type"))
    
    Set M_Objrs = Nothing
    
    Call Cari_LPD_LPA_Payment_2
    
    With LvPTP.SelectedItem
    
        If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
            If CmbTampilkan.text = "PTP DISC." And _
               LvPTP.SelectedItem.SubItems(4) = "Belum di Approve" Then
                FrmViewPTP.cmdapprove.Caption = "Cetak"
            Else
                FrmViewPTP.cmdapprove.Caption = "Approve"
            End If
        End If
           
        FrmViewPTP.CmbJenisPTP.text = Trim(IIf(IsNull(.SubItems(1)), "", .SubItems(1)))
        
        FrmViewPTP.CmbPaymentHandle.text = IIf(IsNull(.SubItems(31)), "", .SubItems(31))
        FrmViewPTP.CmbOccupation.text = IIf(IsNull(.SubItems(32)), "", .SubItems(32))
        FrmViewPTP.CmbReason.text = IIf(IsNull(.SubItems(33)), "", .SubItems(33))
        
        FrmViewPTP.txtothers.text = IIf(IsNull(.SubItems(30)), "", .SubItems(30))
        
        FrmViewPTP.txtproduct.text = TypeAcc
        
        
        SqlWaktu = "select now()"
        Set m_waktuserver = New ADODB.Recordset
        m_waktuserver.CursorLocation = adUseClient
        m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        FrmViewPTP.dtpropsal.Value = m_waktuserver(0)
        
        FrmViewPTP.TxtIdCpa.text = IIf(IsNull(.text), "", .text)
        FrmViewPTP.txtcardno.text = IIf(IsNull(.SubItems(2)), "", .SubItems(2))
        FrmViewPTP.TxtName.text = IIf(IsNull(.SubItems(3)), "", .SubItems(3))
        FrmViewPTP.lblLastPay.Value = IIf(IsNull(.SubItems(7)), "0", Replace(.SubItems(7), ",", ""))
        FrmViewPTP.tdbisnstallment.Value = IIf(IsNull(.SubItems(8)), "1", .SubItems(8))
        
        FrmViewPTP.txtprincipal.Value = IIf(IsNull(.SubItems(11)), "0", Replace(.SubItems(11), ",", ""))
        FrmViewPTP.Label8.text = IIf(IsNull(.SubItems(11)), "0", Replace(.SubItems(11), ",", ""))
        
        FrmViewPTP.txtbalance.Value = IIf(IsNull(.SubItems(12)), "0", Replace(.SubItems(12), ",", ""))
        FrmViewPTP.Label5.text = IIf(IsNull(.SubItems(12)), "0", Replace(.SubItems(12), ",", ""))
        
        FrmViewPTP.txtcharge.Value = IIf(IsNull(.SubItems(17)), "0", Replace(.SubItems(17), ",", ""))
        FrmViewPTP.txtjust.text = IIf(IsNull(.SubItems(21)), "", .SubItems(21))
        FrmViewPTP.txtdownpayment.Value = IIf(IsNull(.SubItems(16)), "0", Replace(.SubItems(16), ",", ""))
        
        If .SubItems(22) = "1" Then
            FrmViewPTP.chkfaxed.Value = 1
        End If
        
        If .SubItems(23) = "1" Then
            FrmViewPTP.chkwentalk.Value = 1
        End If
        
        If .SubItems(24) = "1" Then
            FrmViewPTP.chkKTP.Value = 1
        End If
        
        If .SubItems(25) = "1" Then
            FrmViewPTP.chkpp.Value = 1
        End If
        
        If .SubItems(26) = "1" Then
            FrmViewPTP.chkbillings.Value = 1
        End If
        
        If .SubItems(27) = "1" Then
            FrmViewPTP.Check1.Value = 1
        End If
        
        FrmViewPTP.txtcollect.text = .SubItems(28)
        '@@20062012 Tambahan DOb
        FrmViewPTP.TxtDob.text = IIf(.SubItems(29) = "", "", .SubItems(29))
        'FrmViewPTP.txtproduct.Text = "CARD"
        FrmViewPTP.txtplace.text = "CardHolder"
        FrmViewPTP.Show vbModal
    End With
End Sub

Private Sub Cari_LPD_LPA_Payment_2()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmdsql = "select paydate,payment from tbllunas where custid='"
    cmdsql = cmdsql + Trim(LvPTP.SelectedItem.SubItems(2)) + "' order by paydate desc limit 1 "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        With FrmViewPTP
            If M_Objrs.RecordCount > 0 Then
                .TxtLPDPayment.text = IIf(IsNull(M_Objrs("paydate")), "", Format(M_Objrs("paydate"), "yyyy-mm-dd"))
                .TxtLPAPayment.Value = IIf(IsNull(M_Objrs("payment")), "0", M_Objrs("payment"))
                LpdPayment = "'" + TxtLPDPayment.text + "'"
            Else
                LpdPayment = "null"
                TxtLPDPayment = ""
                .TxtLPAPayment.Value = "0"
            End If
        End With
    Set M_Objrs = Nothing
End Sub

Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function


Private Sub TxtCustid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCari_Click
    End If
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCari_Click
    End If
End Sub
'----------------------------221012 App By VP ----------------------------------------------------------
Private Sub BikinCPA_AppVP(K As Integer)
    Dim cmdsql As String
    Dim Remarks As String
    Dim M_Objrs_Cek_Type As ADODB.Recordset
    Dim TypeAcc As String
    
    TypeAcc = ""

    '@@13022013 Cek type account dulu nih .. pil/card
    cmdsql = "select acc_type from mgm where custid='"
    cmdsql = cmdsql & CStr(LvPTP.ListItems(K).SubItems(2)) & "'"
    Set M_Objrs_Cek_Type = New ADODB.Recordset
    M_Objrs_Cek_Type.CursorLocation = adUseClient
    M_Objrs_Cek_Type.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek_Type.RecordCount > 0 Then
        TypeAcc = IIf(IsNull(M_Objrs_Cek_Type("acc_type")), "", M_Objrs_Cek_Type("acc_type"))
    End If
    
    Set M_Objrs_Cek_Type = Nothing
    
    
    
    Call Cari_LPD_LPA_Payment(K)
    
    cmdsql = "insert into tblcpa (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
    cmdsql = cmdsql + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
    cmdsql = cmdsql + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
    cmdsql = cmdsql + "chksup,chkbillings,chkothers,lpd_from_payment,lpa_from_payment,"
    cmdsql = cmdsql + "f_system,dob,status_ptp,ketother "
    
    '@@19062012 Jika Status PTP DISCON Catat Approvenya
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = cmdsql + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    'Catet Juga yang PTP No Discon 20062012
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = cmdsql + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    '@@16-07-2012 Buat Catet Payment Handle
    cmdsql = cmdsql + " ,vpaymenthandle,voccupation,vreason "
    
    cmdsql = cmdsql + ") values ("
    cmdsql = cmdsql + "now(),'"
    cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
    cmdsql = cmdsql + TypeAcc + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(15)), "0", Replace(LvPTP.ListItems(K).SubItems(15), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(16)), "0", Replace(LvPTP.ListItems(K).SubItems(16), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(17)), "0", Replace(LvPTP.ListItems(K).SubItems(17), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(18)), "0", Replace(LvPTP.ListItems(K).SubItems(18), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "',"
    cmdsql = cmdsql + "now(),'"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(3)), "", LvPTP.ListItems(K).SubItems(3))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(21)), "", LvPTP.ListItems(K).SubItems(21))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "0", Replace(LvPTP.ListItems(K).SubItems(12), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(11)), "0", Replace(LvPTP.ListItems(K).SubItems(11), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(22)), "", LvPTP.ListItems(K).SubItems(22))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(23)), "", LvPTP.ListItems(K).SubItems(23))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(24)), "", LvPTP.ListItems(K).SubItems(24))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(25)), "", LvPTP.ListItems(K).SubItems(25))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(26)), "", LvPTP.ListItems(K).SubItems(26))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(27)), "", LvPTP.ListItems(K).SubItems(27))) + "',"
    cmdsql = cmdsql + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
    cmdsql = cmdsql + CStr(TxtLPAPayment.Value) + "','1',"
    '@@20062012 Tambahkan DOB dan Status PTP
    cmdsql = cmdsql + IIf(LvPTP.ListItems(K).SubItems(29) = "", "null", "'" + LvPTP.ListItems(K).SubItems(29) + "'")
    cmdsql = cmdsql + ",'" + LvPTP.ListItems(K).SubItems(1) + "',' "
    '@@21062012 Tambahkan Keterangan Other
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(30)), "", LvPTP.ListItems(K).SubItems(30)) + "' "
    
    '@@19062012 Buat nyatet approvenya
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = cmdsql + ",now(),'1','"
        cmdsql = cmdsql + Trim(cmbapprove.text) + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "'"
     End If
     
     'Buat nyatet yang jenisnya PTP NO Discount.
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = cmdsql + ",now(),'1','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "'"
     End If
    
    cmdsql = cmdsql + ",'"
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(31)), "", LvPTP.ListItems(K).SubItems(31)) + "','"
    
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(32)), "", LvPTP.ListItems(K).SubItems(32)) + "','"
    cmdsql = cmdsql + IIf(IsNull(LvPTP.ListItems(K).SubItems(33)), "", LvPTP.ListItems(K).SubItems(33)) + "')"
    DoEvents
    M_OBJCONN.execute cmdsql
    
    '@@19062012 Bikin Remarks untuk CPA
     '@@11092012 Tulis Remarks baik untuk yang ptp discon/no discon
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        Remarks = "PtpDisc-"
     Else
        Remarks = "PTPNoDisc-"
     End If
        Remarks = Remarks + "CPA Pengajuan Ke:" + "Pak Hamanto " + "-"
        Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(7)), "", LvPTP.ListItems(K).SubItems(7))) + " -"
        Remarks = Remarks + "Instl: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + " -"
        Remarks = Remarks + "From Bal.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "", LvPTP.ListItems(K).SubItems(12))) + " -"
        Remarks = Remarks + "From Prin.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(14)), "", LvPTP.ListItems(K).SubItems(14))) + " -"
        Remarks = Remarks + "%Balance: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "% -"
        Remarks = Remarks + "%Principal: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "% #USER LOG:" + MDIForm1.Text1.text
        
        cmdsql = "insert into mgm_hst (custid, agent, products, "
        cmdsql = cmdsql + "hst,user_log) values ('"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
        cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(28) + "','"
        cmdsql = cmdsql + "Collection" + "','"
        cmdsql = cmdsql + Remarks + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "')"
        M_OBJCONN.execute cmdsql
    
    
    '@@25072012,Update yang approve dan tanggal proposalnya di tabel tblsendptp jika PTP discount
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='To Be Approved By Pak Hamanto',"
        cmdsql = cmdsql + " log_approve='"
        'CMDSQL = CMDSQL + CStr(Trim(CmbApprove.Text)) + "', log_approve='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "', sts_app_vp='1' "
        cmdsql = cmdsql + " where id='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.execute cmdsql
    End If
    
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "', log_approve='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "', sts_app_vp='1' "
        cmdsql = cmdsql + " where id='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.execute cmdsql
    End If
End Sub

Private Sub KirimPesan_AppVP(K As Integer)
    Dim cmdsql As String
    Dim Remarks As String
    Dim M_Objrs As ADODB.Recordset
    
    Remarks = "Pembuatan PTP untuk custid: " & LvPTP.ListItems(K).SubItems(2) & " sedang dalam proses pengajuan ke Pak Hamanto!"
    
    cmdsql = "insert into msgtbl "
    cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
    cmdsql = cmdsql + LvPTP.ListItems(K).SubItems(28) + "','"
    cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
    cmdsql = cmdsql + MDIForm1.Text1.text + "','"
    cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    cmdsql = cmdsql + Remarks + "')"
    M_OBJCONN.execute cmdsql
        
        
    '@@19072012 Kirim Pesan Buat Ke TL
    'Cari Nama TLNYA
    cmdsql = "select team from usertbl where userid='"
    cmdsql = cmdsql + CStr(Trim(LvPTP.ListItems(K).SubItems(28))) + "' "
    cmdsql = cmdsql + " and team is not null "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        cmdsql = "insert into msgtbl "
        cmdsql = cmdsql + "(recipient, datetime, sender, sentfrom, msg) values ('"
        cmdsql = cmdsql + CStr(Trim(M_Objrs("team"))) + "','"
        cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        cmdsql = cmdsql + Remarks + "')"
        M_OBJCONN.execute cmdsql
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub headerlistappptp()
    ListView1.ColumnHeaders.clear
    With ListView1.ColumnHeaders
        .ADD 1, , "CUST ID", 1000
        .ADD 2, , "ID", 0
        .ADD 3, , "NAMA", 1500
        .ADD 4, , "PROMISE DATE", 1500
        .ADD 5, , "PROMISE PAY", 1500
        .ADD 6, , "STATUSPTP", 0
        .ADD 7, , "VIA", 0
        .ADD 8, , "AGENT", 1000
        .ADD 9, , "TEAM", 1000
        .ADD 10, , "REQ DATE", 1000
        .ADD 11, , "LPD", 1500
        .ADD 12, , "LPA", 1500
        .ADD 13, , "TANGGAL TAGIH", 0
        .ADD 14, , "TENOR", 0
    End With
End Sub


Private Sub headerlistappptp_log()
    ListView1.ColumnHeaders.clear
    With ListView1.ColumnHeaders
        .ADD 1, , "CUST ID", 1000
        .ADD 2, , "ID", 0
        .ADD 3, , "PROMISE DATE", 1500
        .ADD 4, , "PROMISE PAY", 1500
        .ADD 5, , "STATUSPTP", 0
        .ADD 6, , "VIA", 0
        .ADD 7, , "AGENT", 1000
        .ADD 8, , "REQ DATE", 1000
        .ADD 9, , "TANGGAL TAGIH", 0
        .ADD 10, , "TENOR", 0
        .ADD 11, , "APP BY", 1000
        .ADD 12, , "APP DATE", 1000
        .ADD 13, , "STATUS", 1000
    End With
End Sub


'@@221012 Buat Approve Hamanto ---------------------------------------------------------------------------
Private Sub HeaderAppHamanto()
    LvPTP.ColumnHeaders.clear
    With LvHamanto.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Nama CH", 3000
        .ADD 5, , "Status", 2000
        .ADD 6, , "Tanggal Approve", 2000
        .ADD 7, , "Tgl.Payment Effective", 2500
        .ADD 8, , "Total Amount", 1000
        .ADD 9, , "Tenor", 700
        .ADD 10, , "Pembayaran Via", 2000
        .ADD 11, , "Tgl.Tagih", 1500
        .ADD 12, , "Principal", 1000
        .ADD 13, , "Balance", 1000
        .ADD 14, , "Pembayaran Awal", 2000
        .ADD 15, , "Principal", 2000
        .ADD 16, , "Total Payment", 2000
        .ADD 17, , "Down Payment", 2000
        .ADD 18, , "Charge", 2000
        .ADD 19, , "Discount", 2000
        .ADD 20, , "From o/s balance %", 2000
        .ADD 21, , "Principal %", 2000
        .ADD 22, , "Justtification", 2000
        .ADD 23, , "Fax", 800
        .ADD 24, , "When Talking Surlun", 800
        .ADD 25, , "KTP", 800
        .ADD 26, , "Surper", 800
        .ADD 27, , "Billing", 800
        .ADD 28, , "Other", 800
        .ADD 29, , "Agent", 800
        .ADD 30, , "DOB", 1000
        .ADD 31, , "Ket.Other", 1000
        
        '@@ 16-07-2012 Tambahan Payment Handle
        .ADD 32, , "Payment Handle", 2000
        
        '@@17-07-2012 Tambahan Occupation dan Reason
        .ADD 33, , "Occupation", 2000
        .ADD 34, , "Reason", 2000
'
'        .ADD 35, , "Segment", 800
'        .ADD 36, , "Total App", 800
    End With
End Sub

Private Sub BikinCPA_Hamanto(K As Integer)
    Dim cmdsql As String
    Dim Remarks As String
    Dim M_Objrs_Cek_Type As ADODB.Recordset
    Dim TypeAcc As String
    
    TypeAcc = ""

    '@@13022013 Cek type account dulu nih .. pil/card
    cmdsql = "select acc_type from mgm where custid='"
    cmdsql = cmdsql & CStr(LvPTP.ListItems(K).SubItems(2)) & "'"
    Set M_Objrs_Cek_Type = New ADODB.Recordset
    M_Objrs_Cek_Type.CursorLocation = adUseClient
    M_Objrs_Cek_Type.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek_Type.RecordCount > 0 Then
        TypeAcc = IIf(IsNull(M_Objrs_Cek_Type("acc_type")), "", M_Objrs_Cek_Type("acc_type"))
    End If
    
    Set M_Objrs_Cek_Type = Nothing
    
    
    Call Cari_LPD_LPA_Payment_Hamanto(K)
    
    
    cmdsql = "insert into tblcpa (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
    cmdsql = cmdsql + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
    cmdsql = cmdsql + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
    cmdsql = cmdsql + "chksup,chkbillings,chkothers,lpd_from_payment,lpa_from_payment,"
    cmdsql = cmdsql + "f_system,dob,status_ptp,ketother "
    
    '@@19062012 Jika Status PTP DISCON Catat Approvenya
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = cmdsql + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    'Catet Juga yang PTP No Discon 20062012
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = cmdsql + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    '@@16-07-2012 Buat Catet Payment Handle
    cmdsql = cmdsql + " ,vpaymenthandle,voccupation,vreason "
    
    cmdsql = cmdsql + ") values ('"
    'Cmdsql = Cmdsql + "now(),'"
    cmdsql = cmdsql + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "','"
    
    cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "','"
    cmdsql = cmdsql + TypeAcc + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(15)), "0", Replace(LvHamanto.ListItems(K).SubItems(15), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(16)), "0", Replace(LvHamanto.ListItems(K).SubItems(16), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(17)), "0", Replace(LvHamanto.ListItems(K).SubItems(17), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(18)), "0", Replace(LvHamanto.ListItems(K).SubItems(18), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(19)), "", LvHamanto.ListItems(K).SubItems(19))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(20)), "", LvHamanto.ListItems(K).SubItems(20))) + "',"
    cmdsql = cmdsql + "now(),'"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(3)), "", LvHamanto.ListItems(K).SubItems(3))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(21)), "", LvHamanto.ListItems(K).SubItems(21))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(12)), "0", Replace(LvHamanto.ListItems(K).SubItems(12), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(11)), "0", Replace(LvHamanto.ListItems(K).SubItems(11), ",", ""))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(8)), "", LvHamanto.ListItems(K).SubItems(8))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(22)), "", LvHamanto.ListItems(K).SubItems(22))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(23)), "", LvHamanto.ListItems(K).SubItems(23))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(24)), "", LvHamanto.ListItems(K).SubItems(24))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(25)), "", LvHamanto.ListItems(K).SubItems(25))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(26)), "", LvHamanto.ListItems(K).SubItems(26))) + "','"
    cmdsql = cmdsql + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(27)), "", LvHamanto.ListItems(K).SubItems(27))) + "',"
    cmdsql = cmdsql + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
    cmdsql = cmdsql + CStr(TxtLPAPayment.Value) + "','1',"
    '@@20062012 Tambahkan DOB dan Status PTP
    cmdsql = cmdsql + IIf(LvHamanto.ListItems(K).SubItems(29) = "", "null", "'" + LvHamanto.ListItems(K).SubItems(29) + "'")
    cmdsql = cmdsql + ",'" + LvHamanto.ListItems(K).SubItems(1) + "',' "
    '@@21062012 Tambahkan Keterangan Other
    cmdsql = cmdsql + IIf(IsNull(LvHamanto.ListItems(K).SubItems(30)), "", LvHamanto.ListItems(K).SubItems(30)) + "' "
    
    '@@19062012 Buat nyatet approvenya
     If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        cmdsql = cmdsql + ",'" + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "','1','"
        'Cmdsql = Cmdsql + ",now(),'1','"
        'CMDSQL = CMDSQL + Trim(CmbApprove.Text) + "','"
        cmdsql = cmdsql + "Hamanto" + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "'"
     End If
     
     'Buat nyatet yang jenisnya PTP NO Discount.
     If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        cmdsql = cmdsql + ",'" + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "','1','"
        'Cmdsql = Cmdsql + ",now(),'1','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "'"
     End If
    
    cmdsql = cmdsql + ",'"
    cmdsql = cmdsql + IIf(IsNull(LvHamanto.ListItems(K).SubItems(31)), "", LvHamanto.ListItems(K).SubItems(31)) + "','"
    
    cmdsql = cmdsql + IIf(IsNull(LvHamanto.ListItems(K).SubItems(32)), "", LvHamanto.ListItems(K).SubItems(32)) + "','"
    cmdsql = cmdsql + IIf(IsNull(LvHamanto.ListItems(K).SubItems(33)), "", LvHamanto.ListItems(K).SubItems(33)) + "')"
    DoEvents
    M_OBJCONN.execute cmdsql
    
    '@@19062012 Bikin Remarks untuk CPA
     '@@11092012 Tulis Remarks baik untuk yang ptp discon/no discon
     If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        Remarks = "PtpDisc-"
     Else
        Remarks = "PTPNoDisc-"
     End If
        Remarks = Remarks + "App By:" + "Pak Hamanto" + "-"
        Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(7)), "", LvHamanto.ListItems(K).SubItems(7))) + " -"
        Remarks = Remarks + "Instl: " + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(8)), "", LvHamanto.ListItems(K).SubItems(8))) + " -"
        Remarks = Remarks + "From Bal.: Rp." + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(12)), "", LvHamanto.ListItems(K).SubItems(12))) + " -"
        Remarks = Remarks + "From Prin.: Rp." + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(14)), "", LvHamanto.ListItems(K).SubItems(14))) + " -"
        Remarks = Remarks + "%Balance: " + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(19)), "", LvHamanto.ListItems(K).SubItems(19))) + "% -"
        Remarks = Remarks + "%Principal: " + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(20)), "", LvHamanto.ListItems(K).SubItems(20))) + "% #USER LOG:" + MDIForm1.Text1.text
        
        cmdsql = "insert into mgm_hst (custid, agent, products, "
        cmdsql = cmdsql + "hst,user_log) values ('"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "','"
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(28) + "','"
        cmdsql = cmdsql + "Collection" + "','"
        cmdsql = cmdsql + Remarks + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "')"
        M_OBJCONN.execute cmdsql
    
    
    '@@25072012,Update yang approve dan tanggal proposalnya di tabel tblsendptp jika PTP discount
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        'Cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='"
        cmdsql = "update tblsendptp set approve_by='Hamanto', log_approve='"
        'CMDSQL = CMDSQL + CStr(Trim(CmbApprove.Text)) + "', log_approve='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "',tgl_approve_vp='"
        cmdsql = cmdsql + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "',tgl_proposal='"
        cmdsql = cmdsql + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "' "
        cmdsql = cmdsql + " where id='"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).text) + "'"
        M_OBJCONN.execute cmdsql
    End If
    
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        'Cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='"
        cmdsql = "update tblsendptp set approve_by='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "', log_approve='"
        cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.text)) + "',tgl_approve_vp='"
        cmdsql = cmdsql + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "',tgl_proposal=now() "
        cmdsql = cmdsql + " where id='"
        cmdsql = cmdsql + CStr(LvPTP.ListItems(K).text) + "' "
        M_OBJCONN.execute cmdsql
    End If
End Sub

Private Sub Cari_LPD_LPA_Payment_Hamanto(K As Integer)
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    StatusPTP = ""
    TxtLPDPayment.text = ""
    TxtLPAPayment.Value = "0"
    
    cmdsql = "select paydate,payment from tbllunas where custid='"
    cmdsql = cmdsql + Trim(LvHamanto.ListItems(K).SubItems(2)) + "' order by paydate desc limit 1 "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs.RecordCount > 0 Then
            StatusPTP = "PTP-POP"
            TxtLPDPayment.text = IIf(IsNull(M_Objrs("paydate")), "", Format(M_Objrs("paydate"), "yyyy-mm-dd"))
            TxtLPAPayment.Value = IIf(IsNull(M_Objrs("payment")), "0", M_Objrs("payment"))
            LpdPayment = "'" + TxtLPDPayment.text + "'"
        Else
            StatusPTP = "PTP-NEW"
            'LpdPayment = "null"
            TxtLPDPayment.text = ""
            TxtLPAPayment.Value = "0"
        End If
    Set M_Objrs = Nothing
End Sub


Private Sub BikinPTP_Hamanto(K As Integer)
    Dim cmdsql As String
    Dim i As Integer
    Dim M_Objrs_Cek_Tgl As ADODB.Recordset
    
    
    bcekptp = True
    
        'Jika Tenor=1
        If Val(LvHamanto.ListItems(K).SubItems(8)) = 1 Then
                  
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp where custid='"
                cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(6)) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                  
            jatuhtempo = LvHamanto.ListItems(K).SubItems(6)
            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.execute cmdsql
            
            
            ' isi ke tbl log_ptp
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'" + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.execute cmdsql
                
        Else
            'Untuk Tenor yang lebih dari 1
                        
                'Hapus Reserved Data
                cmdsql = "delete from tblreserve where custid='"
                cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
                M_OBJCONN.execute cmdsql
                        
                jatuhtempo = CStr(LvHamanto.ListItems(K).SubItems(6))
            
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp where custid='"
                cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.execute cmdsql
            
            
            'isi ke tbl log_ptp
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            cmdsql = cmdsql + "now(), "
            cmdsql = cmdsql + "'" + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.execute cmdsql
                
                
            n = 0
            
            Call HitungInstallmentPtp_Hamanto(K)
            
            For i = 1 To (Val(LvHamanto.ListItems(K).SubItems(8)) - 1)
                    n = n + 1
                    'JMLPAY = ((.TxtPayment - txtPembayaranAwal.Value) - PaymentTenor) / (.txttenor.Value - 1)
                    JmlPay = PaymentTenor
                    Vrdate = DateAdd("m", n, Format(LvHamanto.ListItems(K).SubItems(6), "yyyy-mm-dd"))
                    
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblreserve where custid='"
                cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblreserve where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                    
                    cmdsql = "INSERT INTO tblreserve "
                    cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                    cmdsql = cmdsql + "VALUES "
                    cmdsql = cmdsql + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
                    cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
                    cmdsql = cmdsql + "now(), "
                    cmdsql = cmdsql + "'IPO')"
                    M_OBJCONN.execute cmdsql
                    

                    
                    cmdsql = "INSERT INTO TblNegoptp_log "
                    cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
                    cmdsql = cmdsql + "VALUES "
                    cmdsql = cmdsql + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
                    cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
                    cmdsql = cmdsql + "now(), "
                    cmdsql = cmdsql + "'" + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','R')"
                    M_OBJCONN.execute cmdsql
        

            Next i
       End If
    PaymentTenor = 0
End Sub

Private Sub HitungInstallmentPtp_Hamanto(K As Integer)
    Dim installment As Double
    
        If Val(LvHamanto.ListItems(K).SubItems(8)) = 0 Or Val(LvHamanto.ListItems(K).SubItems(8)) = 1 Then
            installment = Val(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) / 1
        Else
            installment = (Val(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) - Val(Replace(LvHamanto.ListItems(K).SubItems(13), ",", ""))) / (Val(LvHamanto.ListItems(K).SubItems(8)) - 1)
        End If
        PaymentTenor = Ceiling(installment)
End Sub


Private Sub CatetLogApprove_Hamanto(K As Integer)
    Dim cmdsql As String
        
    cmdsql = "insert into tblsendptp_log_approve "
    cmdsql = cmdsql + "select * from tblsendptp where id='"
    cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).text) + "'"
    DoEvents
    M_OBJCONN.execute cmdsql
End Sub

Private Sub BikinStatusPTP_Hamanto(K As Integer)
    Dim cmdsql As String
    Dim Cmdsql_Cek As String
    Dim StatusRemarks As String
    Dim M_Objrs_Cek As ADODB.Recordset
    Dim AmountNew As Double
    
    AmountNew = 0
    
    Cmdsql_Cek = "select * from tblnegoptp where custid='"
    Cmdsql_Cek = Cmdsql_Cek + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' order by id desc limit 1"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Cek.RecordCount > 0 Then
        AmountNew = Val(IIf(IsNull(M_Objrs_Cek("promisepay")), "0", M_Objrs_Cek("promisepay")))
    Else
       AmountNew = 0
    End If
    
    'Jika StatusPTP=PTP NEW
    If StatusPTP = "PTP-NEW" Then
        Dim M_Objrs_Cek_Status As ADODB.Recordset
        Dim Cmdsql_Cek_status As String
        Dim TglPTPNew As String
        
        'Cari apakah sebelumnya status data=ptp new, jika iya maka tglptpnew tidak usah diupdate
        'Tapi jika status sebelumnya bukan ptp new maka update tglptpnew=now
        Cmdsql_Cek_status = "select * from mgm where custid='"
        Cmdsql_Cek_status = Cmdsql_Cek_status + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
        Set M_Objrs_Cek_Status = New ADODB.Recordset
        M_Objrs_Cek_Status.CursorLocation = adUseClient
        M_Objrs_Cek_Status.Open Cmdsql_Cek_status, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_Status.RecordCount > 0 Then
            If M_Objrs_Cek_Status("tglptpnew") = "" Or IsNull(M_Objrs_Cek_Status("tglptpnew")) = True _
               Or M_Objrs_Cek_Status("tglptpnew") = Empty Then
                TglPTPNew = "now()"
             Else
                TglPTPNew = "'" + CStr(Format(M_Objrs_Cek_Status("tglptpnew"), "yyyy-mm-dd")) + "'"
             End If
        End If
        
        Set M_Objrs_Cek_Status = Nothing
    
        cmdsql = "update mgm set dateptpnew='"
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(6) + "',tgl_tagih='"
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(10) + "', amountnew='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',tglallptp='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + "',tglallptp='"
        
        '@@20062012, amountnew ambil dari negoptp terakhir aja deh....
        cmdsql = cmdsql + CStr(AmountNew) + "',tglallptp='"
        
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(6) + "',f_cek_new='PTP-NE',"
        cmdsql = cmdsql + "tglincoming=now(),ttlptp='"
        cmdsql = cmdsql + CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) + "',"
        cmdsql = cmdsql + "kethslkerja_new='PTP-NEW',kethslkerjadesc_new='PTP-NEW',ptpvia='"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-NEW', dateptp='"
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(6) + "',tglptpnew=" + TglPTPNew
        cmdsql = cmdsql + ",tenor='"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(8)) + "' "
        cmdsql = cmdsql + "where custid='"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
        DoEvents
        M_OBJCONN.execute cmdsql
        
    End If
    
    If StatusPTP = "PTP-POP" Then
        cmdsql = "update mgm set dateptp='"
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(6) + "',tgl_tagih='"
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(10) + "',tglallptp='"
        cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(6) + "',f_cek_new='PTP-PO',"
        cmdsql = cmdsql + "tglincoming=now(),ttlptp='"
        cmdsql = cmdsql + CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) + "',"
        cmdsql = cmdsql + "kethslkerja_new='PTP-POP',kethslkerjadesc_new='PTP-POP',ptpvia='"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-POP',amountptp='"
        cmdsql = cmdsql + CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) + "',tenor='"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(8)) + "' "
        cmdsql = cmdsql + "where custid='"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
        M_OBJCONN.execute cmdsql
    End If
    
     '@@19062012 Bikin Remark Status PTP
        StatusRemarks = "PTP Approve by: " & MDIForm1.Text1.text & "/"
        StatusRemarks = StatusRemarks & "Jenis PTP:" & StatusPTP & "/"
        StatusRemarks = StatusRemarks & "Amount PTP:"
        StatusRemarks = StatusRemarks & CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "PTP Via:" & ""
        StatusRemarks = StatusRemarks & CStr(Replace(LvHamanto.ListItems(K).SubItems(9), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "Date PTP:" & Format(LvHamanto.ListItems(K).SubItems(6), "yyyy-mm-dd")
        
        cmdsql = "insert into mgm_hst(custid,agent,hst,f_cek_new,user_log) values ('"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(2)) + "','"
        cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','"
        cmdsql = cmdsql + StatusRemarks + "','"
        cmdsql = cmdsql + StatusPTP + "','"
        cmdsql = cmdsql + CStr(MDIForm1.Text1.text) + "')"
        M_OBJCONN.execute cmdsql
        
   Set M_Objrs_Cek = Nothing
        
End Sub

Private Sub HapusData_Hamanto(K As Integer)
    Dim cmdsql As String
    
    cmdsql = "delete from tblsendptp where id='"
    cmdsql = cmdsql + CStr(LvHamanto.ListItems(K).text) + "'"
    M_OBJCONN.execute cmdsql
End Sub

Private Sub KirimPesan_Hamanto(K As Integer)
    Dim cmdsql As String
    Dim Remarks As String
    Dim M_Objrs As ADODB.Recordset
    
    Remarks = "Pembuatan PTP untuk custid: " & LvHamanto.ListItems(K).SubItems(2) & " telah di approve!"
    
    cmdsql = "insert into msgtbl "
    cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
    cmdsql = cmdsql + LvHamanto.ListItems(K).SubItems(28) + "','"
    cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
    cmdsql = cmdsql + MDIForm1.Text1.text + "','"
    cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    cmdsql = cmdsql + Remarks + "')"
    M_OBJCONN.execute cmdsql
        
        
    '@@19072012 Kirim Pesan Buat Ke TL
    'Cari Nama TLNYA
    cmdsql = "select team from usertbl where userid='"
    cmdsql = cmdsql + CStr(Trim(LvHamanto.ListItems(K).SubItems(28))) + "' "
    cmdsql = cmdsql + " and team is not null "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        cmdsql = "insert into msgtbl "
        cmdsql = cmdsql + "(recipient, datetime, sender, sentfrom, msg) values ('"
        cmdsql = cmdsql + CStr(Trim(M_Objrs("team"))) + "','"
        cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        cmdsql = cmdsql + Remarks + "')"
        M_OBJCONN.execute cmdsql
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub My_Export_Excel()
    Dim a           As Long
    Dim b           As Long
    Dim ExlObj      As Excel.Application
    Dim listcustid  As String
    Dim rs          As ADODB.Recordset
    Dim iRow        As Integer
    Dim i           As Integer
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            listcustid = listcustid & ",'" & LvPTP.ListItems(K).SubItems(2) & "'"
        End If
    Next K
    
    listcustid = Mid(listcustid, 2)
    
    'Strsql = "select custid,vcustname,'" & CmbApprove.Text & "' as Approved,* FROM tblsendptp WHERE custid in (" & listcustid & ")"
'    Strsql = "SELECT * FROM ("
'    Strsql = Strsql + " SELECT custid,vcustname,'" & CmbApprove.Text & "' as Approved,* "
'    Strsql = Strsql + " FROM tblsendptp WHERE custid in (" & listcustid & ")) As a"
'    Strsql = Strsql + " LEFT JOIN (SELECT custid, OpenDate, b_d, FROM mgm WHERE custid in (" & listcustid & ")) As b"
'    Strsql = Strsql + " on a.custid = b.custid"
    
    Strsql = "SELECT * FROM ("
    Strsql = Strsql + " SELECT '" & cmbapprove.text & "' as Approved,* "
    Strsql = Strsql + " FROM tblsendptp WHERE custid in (" & listcustid & ")) As a"
    Strsql = Strsql + " LEFT JOIN (SELECT custid, OpenDate, b_d,LastPay,Pay_Dt FROM mgm WHERE custid in (" & listcustid & ")) As b"
    Strsql = Strsql + " on a.custid = b.custid"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:N1").MergeCells = True
    ExlObj.Range("A2:N2").MergeCells = True
    ExlObj.Range("A4:N4").Font.Bold = True
    
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "List CPA Approve"
        .Cells(1, 1).Font.Name = "Verdana"
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Tanggal : " + Format(Now, "dd-mm-yyyy")
        .Cells(2, 1).Font.Name = "Verdana"
        .Cells(2, 1).Font.Bold = True
        .Cells(4, 1).Value = "NO"
        .Cells(4, 2).Value = "CARD NUMBER"
        .Cells(4, 3).Value = "CH NAME"
        .Cells(4, 4).Value = "APPROVED"
        .Cells(4, 5).Value = "ADMIN CREATED"
        .Cells(4, 6).Value = "RECEIVED BY" 'Dikosongkan
        .Cells(4, 7).Value = "ID"""
        .Cells(4, 8).Value = "Jenis PTP"
        .Cells(4, 9).Value = "Custid"
        .Cells(4, 10).Value = "Nama CH"
        .Cells(4, 11).Value = "Status"
        .Cells(4, 12).Value = "Tanggal Approve"
        .Cells(4, 13).Value = "Tgl.Payment Effective"
        .Cells(4, 14).Value = "Total Amount"
        .Cells(4, 15).Value = "Tenor"
        .Cells(4, 16).Value = "Pembayaran Via"
        .Cells(4, 17).Value = "Tgl.Tagih"
        .Cells(4, 18).Value = "Principal"
        .Cells(4, 19).Value = "Balance"
        .Cells(4, 20).Value = "Pembayaran Awal"
        .Cells(4, 21).Value = "Principal"
        .Cells(4, 22).Value = "Total Payment"
        .Cells(4, 23).Value = "Down Payment"
        .Cells(4, 24).Value = "Charge"
        .Cells(4, 25).Value = "Discount"
        .Cells(4, 26).Value = "From o/s balance %"
        .Cells(4, 27).Value = "Principal %"
        .Cells(4, 28).Value = "Justtification"
        .Cells(4, 29).Value = "Fax"
        .Cells(4, 30).Value = "When Talking Surlun"
        .Cells(4, 31).Value = "KTP"
        .Cells(4, 32).Value = "Surper"
        .Cells(4, 33).Value = "Billing"
        .Cells(4, 34).Value = "Other"
        .Cells(4, 35).Value = "Agent"
        .Cells(4, 36).Value = "DOB"
        .Cells(4, 37).Value = "Ket.Other"
        .Cells(4, 38).Value = "Open Date"
        .Cells(4, 39).Value = "WO Date"
        .Cells(4, 40).Value = "LPD"
        .Cells(4, 41).Value = "LPA"
        
        iRow = 4
        If rs.RecordCount > 0 Then
            PB1.Max = rs.RecordCount
            i = 0
            Do Until rs.EOF
                i = i + 1
                iRow = iRow + 1
                PB1.Value = rs.Bookmark
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 3).Value = IIf(IsNull(rs!vcustname), "", rs!vcustname)
                .Cells(iRow, 4).Value = IIf(IsNull(rs!approved), "", rs!approved)
                .Cells(iRow, 5).Value = MDIForm1.Text1.text
                .Cells(iRow, 6).Value = "" 'Dikosongkan
                .Cells(iRow, 7).Value = ""
                .Cells(iRow, 8).Value = IIf(IsNull(rs!jenis_ptp), "", rs!jenis_ptp)
                .Cells(iRow, 9).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 10).Value = ""
                .Cells(iRow, 11).Value = IIf(IsNull(rs!STATUS), "", rs!STATUS)
                .Cells(iRow, 12).Value = IIf(IsNull(rs("tgl_approve")), "", Format(rs("tgl_approve"), "yyyy-mm-dd"))
                .Cells(iRow, 13).Value = IIf(IsNull(rs("date_payment_effective")), "", Format(rs("date_payment_effective"), "yyyy-mm-dd"))
                .Cells(iRow, 14).Value = IIf(IsNull(rs!total_amount_deal), "", rs!total_amount_deal)
                .Cells(iRow, 15).Value = IIf(IsNull(rs!Tenor), "", rs!Tenor)
                .Cells(iRow, 16).Value = IIf(IsNull(rs!pembayaran_via), "", rs!pembayaran_via)
                .Cells(iRow, 17).Value = IIf(IsNull(rs("tgl_tagih")), "", Format(rs("tgl_tagih"), "yyyy-mm-dd"))
                .Cells(iRow, 18).Value = IIf(IsNull(rs!Principal), "", rs!Principal)
                .Cells(iRow, 19).Value = IIf(IsNull(rs!balance), "", rs!balance)
                .Cells(iRow, 20).Value = IIf(IsNull(rs!Pembayaran_awal), "", rs!Pembayaran_awal)
                .Cells(iRow, 21).Value = IIf(IsNull(rs!Principal), "", rs!Principal)
                .Cells(iRow, 22).Value = IIf(IsNull(rs!nttlpayment), "", rs!nttlpayment)
                .Cells(iRow, 23).Value = IIf(IsNull(rs!ndownpay), "", rs!ndownpay)
                .Cells(iRow, 24).Value = IIf(IsNull(rs!ncharge), "", rs!ncharge)
                .Cells(iRow, 25).Value = IIf(IsNull(rs!ndiscountamt), "", rs!ndiscountamt)
                .Cells(iRow, 26).Value = IIf(IsNull(rs!vosbalance), "", rs!vosbalance)
                .Cells(iRow, 27).Value = IIf(IsNull(rs!vosprincipal), "", rs!vosprincipal)
                .Cells(iRow, 28).Value = IIf(IsNull(rs!vjust), "", rs!vjust)
                .Cells(iRow, 29).Value = IIf(IsNull(rs!chkfaxed), "", rs!chkfaxed)
                .Cells(iRow, 30).Value = IIf(IsNull(rs!chkwentalking), "", rs!chkwentalking)
                .Cells(iRow, 31).Value = IIf(IsNull(rs!chkKTP), "", rs!chkKTP)
                .Cells(iRow, 32).Value = IIf(IsNull(rs!chksup), "", rs!chksup)
                .Cells(iRow, 33).Value = IIf(IsNull(rs!chkbillings), "", rs!chkbillings)
                .Cells(iRow, 34).Value = IIf(IsNull(rs!chkothers), "", rs!chkothers)
                .Cells(iRow, 35).Value = IIf(IsNull(rs!agent), "", rs!agent)
                .Cells(iRow, 36).Value = IIf(IsNull(rs("DOB")), "", Format(rs("DOB"), "yyyy-mm-dd"))
                .Cells(iRow, 37).Value = IIf(IsNull(rs!ket_other), "", rs!ket_other)
'                .Cells(iRow, 38).Value = IIf(IsNull(RS!OpenDate), "", RS!OpenDate)
'                .Cells(iRow, 39).Value = IIf(IsNull(RS!b_d), "", RS!b_d)
                .Cells(iRow, 38).Value = cnull(rs!opendate)
                .Cells(iRow, 39).Value = IIf(IsNull(rs!B_D), "", rs!B_D)
                .Cells(iRow, 40).Value = IIf(IsNull(rs("Pay_Dt")), "", Format(rs("Pay_Dt"), "dd-mm-yyyy"))
                .Cells(iRow, 41).Value = IIf(IsNull(rs!lastpay), "", rs!lastpay)


                rs.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 14
            ExlObj.Cells(4, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
        PB1.Value = 0
        Command1.Enabled = True
    
        Set ExlObj = Nothing
        Set rs = Nothing

        'StartMeUp (Txtlocation.Text)
        'FILL COLOR CELL
        'ExlObj.Range(.Cells(NoUrut, 1), .Cells(NoUrut, 7)).Interior.Color = RGB(6, 207, 250)
    End With
End Sub
