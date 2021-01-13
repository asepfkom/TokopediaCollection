VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSendPTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send PTP"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12690
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Form Send PTP"
      TabPicture(0)   =   "FrmSendPTP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label51"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label12(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(36)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(37)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(38)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(39)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(14)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(15)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(16)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(17)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(19)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label14"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label13"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label12(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label6(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label7(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(24)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label15"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label16"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label17"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label7(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label7(3)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label7(4)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "LblJumlah"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label18"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TxtPaymentMonthSebenarnya"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtPayAfterTenor"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "tdbisnstallment"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtbalance"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtprincipal"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtfuture"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtdownpayment"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lblLastPay"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtdiscount"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtcharge"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtPembayaranAwal"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txttenor"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "TxtTglTagih"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtPayment"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TxtPeymentEffective"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Frame1"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "CmbJenisPTP"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "CmbViaPtp"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "CmdSendPTP"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "CmdBatal"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Check1"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "chkbillings"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "chkpp"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "chkKTP"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "chkwentalk"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "chkfaxed"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "TxtJustification"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtpersenprincipal"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txtfrombalancepersen"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txtperiodpay"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Label5"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Label8"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "TxtDob"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txtothers"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "CmbPaymentHandle"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "CmbOccupation"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "CmbReason"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "ChkPersetujuan"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).ControlCount=   72
      TabCaption(1)   =   "Log Send PTP"
      TabPicture(1)   =   "FrmSendPTP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdUnCekAll"
      Tab(1).Control(1)=   "CmdCekAll"
      Tab(1).Control(2)=   "CmdHapus"
      Tab(1).Control(3)=   "TxtJumlah"
      Tab(1).Control(4)=   "LvLogPTP"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label81"
      Tab(1).ControlCount=   7
      Begin VB.CheckBox ChkPersetujuan 
         Caption         =   "Saya menyatakan telah meminta dokumen kepada customer "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7020
         TabIndex        =   78
         Top             =   7320
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.ComboBox CmbReason 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmSendPTP.frx":0038
         Left            =   2280
         List            =   "FrmSendPTP.frx":004E
         TabIndex        =   77
         Top             =   5220
         Width           =   3015
      End
      Begin VB.ComboBox CmbOccupation 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmSendPTP.frx":00CA
         Left            =   2280
         List            =   "FrmSendPTP.frx":00D7
         TabIndex        =   75
         Top             =   4860
         Width           =   3015
      End
      Begin VB.ComboBox CmbPaymentHandle 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmSendPTP.frx":0109
         Left            =   2280
         List            =   "FrmSendPTP.frx":0119
         TabIndex        =   73
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox txtothers 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   2640
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   7380
         Width           =   3225
      End
      Begin VB.TextBox TxtDob 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7980
         TabIndex        =   68
         Top             =   4020
         Width           =   2175
      End
      Begin VB.TextBox Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2940
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtperiodpay 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7965
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3270
         Width           =   2160
      End
      Begin VB.TextBox txtfrombalancepersen 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7965
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5355
         Width           =   2145
      End
      Begin VB.TextBox txtpersenprincipal 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7965
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   5760
         Width           =   2160
      End
      Begin VB.TextBox TxtJustification 
         BackColor       =   &H00C0FFC0&
         Height          =   975
         Left            =   8940
         TabIndex        =   34
         Top             =   3360
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.CheckBox chkfaxed 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Faxed"
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   5940
         Width           =   1005
      End
      Begin VB.CheckBox chkwentalk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "When Talking Surlun"
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   6180
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox chkKTP 
         BackColor       =   &H00C0FFC0&
         Caption         =   "KTP"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   29
         Top             =   6480
         Width           =   765
      End
      Begin VB.CheckBox chkpp 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Surper"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   28
         Top             =   6480
         Width           =   945
      End
      Begin VB.CheckBox chkbillings 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Billings"
         Enabled         =   0   'False
         Height          =   405
         Left            =   2460
         TabIndex        =   27
         Top             =   6720
         Width           =   825
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Others"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2460
         TabIndex        =   26
         Top             =   7140
         Width           =   795
      End
      Begin VB.CommandButton CmdUnCekAll 
         Caption         =   "&UnCek All"
         Height          =   375
         Left            =   -65220
         TabIndex        =   23
         Top             =   7500
         Width           =   1215
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "&Cek All"
         Height          =   375
         Left            =   -66420
         TabIndex        =   22
         Top             =   7500
         Width           =   1215
      End
      Begin VB.CommandButton CmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   -63660
         TabIndex        =   21
         Top             =   7500
         Width           =   1215
      End
      Begin VB.TextBox TxtJumlah 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73740
         TabIndex        =   20
         Text            =   "0"
         Top             =   7500
         Width           =   1155
      End
      Begin VB.CommandButton CmdBatal 
         Caption         =   "&Batal"
         Height          =   435
         Left            =   8520
         TabIndex        =   16
         Top             =   7620
         Width           =   1515
      End
      Begin VB.CommandButton CmdSendPTP 
         Caption         =   "Send PTP..."
         Height          =   435
         Left            =   6960
         TabIndex        =   15
         Top             =   7620
         Width           =   1515
      End
      Begin VB.ComboBox CmbViaPtp 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmSendPTP.frx":0138
         Left            =   2280
         List            =   "FrmSendPTP.frx":014B
         TabIndex        =   5
         Top             =   4020
         Width           =   3015
      End
      Begin VB.ComboBox CmbJenisPTP 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmSendPTP.frx":017C
         Left            =   2280
         List            =   "FrmSendPTP.frx":0186
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Frame Frame1 
         Caption         =   "Perhatian"
         Height          =   915
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   12375
         Begin VB.Label Label1 
            Caption         =   $"FrmSendPTP.frx":01A9
            Height          =   555
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   300
            Width           =   12015
         End
      End
      Begin TDBDate6Ctl.TDBDate TxtPeymentEffective 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "FrmSendPTP.frx":0258
         Caption         =   "FrmSendPTP.frx":0370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":03DC
         Keys            =   "FrmSendPTP.frx":03FA
         Spin            =   "FrmSendPTP.frx":0458
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
      Begin TDBNumber6Ctl.TDBNumber txtPayment 
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   2280
         Width           =   2250
         _Version        =   65536
         _ExtentX        =   3969
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":0480
         Caption         =   "FrmSendPTP.frx":04A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":050C
         Keys            =   "FrmSendPTP.frx":052A
         Spin            =   "FrmSendPTP.frx":0574
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12648384
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
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBDate6Ctl.TDBDate TxtTglTagih 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   5580
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   503
         Calendar        =   "FrmSendPTP.frx":059C
         Caption         =   "FrmSendPTP.frx":06B4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":0720
         Keys            =   "FrmSendPTP.frx":073E
         Spin            =   "FrmSendPTP.frx":079C
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
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber txttenor 
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   2940
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   441
         Calculator      =   "FrmSendPTP.frx":07C4
         Caption         =   "FrmSendPTP.frx":07E4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":0850
         Keys            =   "FrmSendPTP.frx":086E
         Spin            =   "FrmSendPTP.frx":08B8
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   49152
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin MSComctlLib.ListView LvLogPTP 
         Height          =   6360
         Left            =   -74820
         TabIndex        =   17
         Top             =   900
         Width           =   12180
         _ExtentX        =   21484
         _ExtentY        =   11218
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
      Begin TDBNumber6Ctl.TDBNumber txtPembayaranAwal 
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   2580
         Width           =   2250
         _Version        =   65536
         _ExtentX        =   3969
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":08E0
         Caption         =   "FrmSendPTP.frx":0900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":096C
         Keys            =   "FrmSendPTP.frx":098A
         Spin            =   "FrmSendPTP.frx":09D4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12648384
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
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber txtcharge 
         Height          =   255
         Left            =   7965
         TabIndex        =   37
         Top             =   4680
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":09FC
         Caption         =   "FrmSendPTP.frx":0A1C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":0A88
         Keys            =   "FrmSendPTP.frx":0AA6
         Spin            =   "FrmSendPTP.frx":0AF0
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
      Begin TDBNumber6Ctl.TDBNumber txtdiscount 
         Height          =   255
         Left            =   7965
         TabIndex        =   38
         Top             =   5040
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":0B18
         Caption         =   "FrmSendPTP.frx":0B38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":0BA4
         Keys            =   "FrmSendPTP.frx":0BC2
         Spin            =   "FrmSendPTP.frx":0C0C
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
      Begin TDBNumber6Ctl.TDBNumber lblLastPay 
         Height          =   255
         Left            =   7965
         TabIndex        =   46
         Top             =   2280
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":0C34
         Caption         =   "FrmSendPTP.frx":0C54
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":0CC0
         Keys            =   "FrmSendPTP.frx":0CDE
         Spin            =   "FrmSendPTP.frx":0D28
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
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
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber txtdownpayment 
         Height          =   255
         Left            =   7965
         TabIndex        =   47
         Top             =   2640
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":0D50
         Caption         =   "FrmSendPTP.frx":0D70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":0DDC
         Keys            =   "FrmSendPTP.frx":0DFA
         Spin            =   "FrmSendPTP.frx":0E44
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
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
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber txtfuture 
         Height          =   255
         Left            =   7965
         TabIndex        =   48
         Top             =   2955
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":0E6C
         Caption         =   "FrmSendPTP.frx":0E8C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":0EF8
         Keys            =   "FrmSendPTP.frx":0F16
         Spin            =   "FrmSendPTP.frx":0F60
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
      Begin TDBNumber6Ctl.TDBNumber txtprincipal 
         Height          =   255
         Left            =   7965
         TabIndex        =   49
         Top             =   3675
         Visible         =   0   'False
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":0F88
         Caption         =   "FrmSendPTP.frx":0FA8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":1014
         Keys            =   "FrmSendPTP.frx":1032
         Spin            =   "FrmSendPTP.frx":107C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
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
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber txtbalance 
         Height          =   255
         Left            =   7965
         TabIndex        =   50
         Top             =   1920
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":10A4
         Caption         =   "FrmSendPTP.frx":10C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":1130
         Keys            =   "FrmSendPTP.frx":114E
         Spin            =   "FrmSendPTP.frx":1198
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
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
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber tdbisnstallment 
         Height          =   255
         Left            =   10380
         TabIndex        =   59
         Top             =   3600
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":11C0
         Caption         =   "FrmSendPTP.frx":11E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":124C
         Keys            =   "FrmSendPTP.frx":126A
         Spin            =   "FrmSendPTP.frx":12B4
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999
         MinValue        =   -999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber TxtPayAfterTenor 
         Height          =   255
         Left            =   2280
         TabIndex        =   63
         Top             =   3300
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":12DC
         Caption         =   "FrmSendPTP.frx":12FC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":1368
         Keys            =   "FrmSendPTP.frx":1386
         Spin            =   "FrmSendPTP.frx":13D0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12648384
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
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber TxtPaymentMonthSebenarnya 
         Height          =   255
         Left            =   2460
         TabIndex        =   65
         Top             =   3660
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calculator      =   "FrmSendPTP.frx":13F8
         Caption         =   "FrmSendPTP.frx":1418
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmSendPTP.frx":1484
         Keys            =   "FrmSendPTP.frx":14A2
         Spin            =   "FrmSendPTP.frx":14EC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   65535
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#########0;;Null"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#########0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   -99999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   7864321
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin VB.Label Label18 
         Caption         =   "Max.250 Karakter"
         Height          =   195
         Left            =   3300
         TabIndex        =   80
         Top             =   8160
         Width           =   1935
      End
      Begin VB.Label LblJumlah 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2640
         TabIndex        =   79
         Top             =   8160
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Reason:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   76
         Top             =   5220
         Width           =   2115
      End
      Begin VB.Label Label7 
         Caption         =   "Occupation:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   74
         Top             =   4860
         Width           =   2115
      End
      Begin VB.Label Label7 
         Caption         =   "Payment Handle By:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   72
         Top             =   4440
         Width           =   2115
      End
      Begin VB.Label Label17 
         Caption         =   "*)Otomatis"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   4200
         TabIndex        =   71
         Top             =   3660
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "*)Otomatis"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   4320
         TabIndex        =   70
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "DOB"
         Height          =   255
         Left            =   6120
         TabIndex        =   67
         Top             =   4020
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment/Month By System:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   3660
         Width           =   3390
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment/Month:"
         Height          =   240
         Index           =   24
         Left            =   120
         TabIndex        =   64
         Top             =   3300
         Width           =   2070
      End
      Begin VB.Label Label7 
         Caption         =   "Principal di database"
         Height          =   285
         Index           =   1
         Left            =   10320
         TabIndex        =   62
         Top             =   2640
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label6 
         Caption         =   "Balance di database"
         Height          =   285
         Index           =   1
         Left            =   10320
         TabIndex        =   61
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Installment Period"
         Height          =   195
         Index           =   1
         Left            =   10320
         TabIndex        =   60
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "::Calculation::"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6180
         TabIndex        =   58
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H003F9E0C&
         Caption         =   "Data CPA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         TabIndex        =   57
         Top             =   1500
         Width           =   6435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Principal"
         Height          =   285
         Index           =   19
         Left            =   6120
         TabIndex        =   56
         Top             =   3720
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment period month"
         Height          =   285
         Index           =   17
         Left            =   6120
         TabIndex        =   55
         Top             =   3315
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Future Payment"
         Height          =   195
         Index           =   16
         Left            =   6120
         TabIndex        =   54
         Top             =   3000
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
         Height          =   195
         Index           =   15
         Left            =   6120
         TabIndex        =   53
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment"
         Height          =   330
         Index           =   14
         Left            =   6120
         TabIndex        =   52
         Top             =   2325
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   285
         Index           =   13
         Left            =   6120
         TabIndex        =   51
         Top             =   1965
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Charge"
         Height          =   240
         Index           =   39
         Left            =   6165
         TabIndex        =   42
         Top             =   4725
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amount"
         Height          =   240
         Index           =   38
         Left            =   6165
         TabIndex        =   41
         Top             =   5085
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From o/s balance %"
         Height          =   330
         Index           =   37
         Left            =   6165
         TabIndex        =   40
         Top             =   5400
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "principal (%) from "
         Height          =   240
         Index           =   36
         Left            =   6165
         TabIndex        =   39
         Top             =   5760
         Width           =   1230
      End
      Begin VB.Label Label12 
         Caption         =   "Justification:"
         Height          =   255
         Index           =   0
         Left            =   6900
         TabIndex        =   33
         Top             =   6300
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Label11 
         Caption         =   "Document:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   5940
         Width           =   2115
      End
      Begin VB.Label Label10 
         Caption         =   "Pembayaran Awal:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2580
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Jumlah Data:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   7560
         Width           =   975
      End
      Begin VB.Label Label81 
         Caption         =   "List Log Send PTP:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "Tanggal Tagih:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   5580
         Width           =   2115
      End
      Begin VB.Label Label6 
         Caption         =   "Pembayaran Via:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   4020
         Width           =   2115
      End
      Begin VB.Label Label51 
         Caption         =   "Date Payment Effective:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "Tenor:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2940
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Total Amount Deal:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Jenis PTP:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1995
      End
   End
End
Attribute VB_Name = "FrmSendPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbJenisPTP_Click()
'    If CmbJenisPTP.Text = "PTP 1X BAYAR" Then
'        txttenor.Value = 1
'        txttenor.Enabled = False
'        txtPembayaranAwal.Enabled = False
'    End If
'    If CmbJenisPTP.Text = "PTP DEAL LUNAS" Then
'        txttenor.Value = 1
'        txttenor.Enabled = True
'        txtPembayaranAwal.Enabled = True
'    End If

'    If CmbJenisPTP.Text = "PTP Discount" Then
'        txtPayment.Value = txtbalance.Value
'        txtPayment.Enabled = False
'    End If
'    If CmbJenisPTP.Text = "PTP No Discount" Then
'        txtPayment.Enabled = True
'    End If
End Sub




Private Sub CmbOccupation_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbPaymentHandle_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub CmbReason_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbViaPtp_Click()
    Call CariTanggalTagih
End Sub

Private Sub CmbViaPtp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvLogPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!"
        Exit Sub
    End If
    
    For W = 1 To LvLogPTP.ListItems.Count
        LvLogPTP.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdHapus_Click()
    Dim a As String
    Dim CMDSQL As String
    Dim K As Integer
    
    If LvLogPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin akan menghapus data yang dicentang?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        Exit Sub
    End If
    
    For K = 1 To LvLogPTP.ListItems.Count
        If LvLogPTP.ListItems(K).Checked = True Then
            CMDSQL = "delete from tblsendptp where id='"
            CMDSQL = CMDSQL + CStr(LvLogPTP.ListItems(K).text) + "' and status='0'"
            M_OBJCONN.execute CMDSQL
        End If
    Next K
    
    MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
    Call isilog
End Sub

Private Sub CmdSendPTP_Click()
    Dim VSAVE As Boolean
    Dim CMDSQL As String
    Dim M_OBJRS As ADODB.Recordset
    Dim Remarks As String
    Dim M_Objrs_Cek As ADODB.Recordset
    Dim W As String
    Dim WK As String
    Dim WA As String
    
    Dim strFaxed As String
    Dim strOthers As String
    Dim strwentalk As String
    Dim strKTP As String
    Dim strSup As String
    Dim strBilling As String
    
    Dim Occupation() As String
    Dim Reason() As String
    
    
    TGL = Format(TxtPeymentEffective.Value, "YYYY-MM-DD")
    
        SqlWaktu = "select now() as tgl"
        Set m_waktuserver = New ADODB.Recordset
        m_waktuserver.CursorLocation = adUseClient
        m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If TGL < Format(m_waktuserver!TGL, "yyyy-mm-dd") Then
        MsgBox "Tidak Boleh Back Date"
        Exit Sub
    End If
    
    VSAVE = True
    VSAVE = VSAVE And CmbJenisPTP.text <> Empty
    VSAVE = VSAVE And TxtPeymentEffective.Value <> Empty
    VSAVE = VSAVE And txtPayment.Value > 0
    VSAVE = VSAVE And txttenor.Value > 0
    VSAVE = VSAVE And CmbViaPtp.text <> Empty
    VSAVE = VSAVE And TxtTglTagih.Value <> Empty
    VSAVE = VSAVE And txtPembayaranAwal.Value <> Empty
    'VSAVE = VSAVE And TxtJustification.Text <> Empty
    VSAVE = VSAVE And CmbPaymentHandle.text <> Empty
    VSAVE = VSAVE And CmbOccupation.text <> Empty
    VSAVE = VSAVE And CmbReason.text <> Empty
    
    If VSAVE = False Then
        MsgBox "Textbox tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    '@@231012,Checkbox Persetujuan dinonaktifkan dulu
'    If ChkPersetujuan.Value = False Then
'        MsgBox "Anda harus mencentang: Saya menyatakan telah meminta dokumen kepada customer!", vbOKOnly + vbInformation, "Informasi!"
'        Exit Sub
'    End If
    
    '18-06-2012 Jika Payment/Month tidak sama dengan Payment Calculasi by system
    'Maka Sistem Payment/Month diubah sama dengan payment calculasi by sistem
    
    If TxtPayAfterTenor.Value <> TxtPaymentMonthSebenarnya.Value Then
        WK = MsgBox("Payment per month anda akan disamakan dengan payment per month hasil kalkulasi sistem! Anda Setuju?", vbYesNo + vbQuestion, "Konfirmasi")
        If WK = vbNo Then
            MsgBox "Send PTP gagal!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        Else
            TxtPayAfterTenor.Value = TxtPaymentMonthSebenarnya.Value
        End If
    End If
    
    
    '@@08062012 Cek data dulu, apakah data pernah diinputkan?
    CMDSQL = "select * from tblsendptp where custid='"
    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Caption) + "' and status='0'"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek.RecordCount > 0 Then
        W = MsgBox("Ada PTP sebelumnya yang belum di approve. Apakah anda akan menghapus data PTP sebelumnya?", vbYesNo + vbQuestion, "Konfirmasi")
        If W = vbYes Then
            While Not M_Objrs_Cek.EOF
                CMDSQL = "delete from tblsendptp where id='"
                CMDSQL = CMDSQL + CStr(M_Objrs_Cek("id")) + "'"
                M_OBJCONN.execute CMDSQL
                M_Objrs_Cek.MoveNext
            Wend
        End If
    End If
    
    Set M_Objrs_Cek = Nothing
    
    '====================== Cek Checkbox ======================================
    strFaxed = ""
    strOthers = ""
    strwentalk = ""
    strKTP = ""
    strSup = ""
    strBilling = ""
    
    StatusChekcBox = ""
    If chkfaxed.Value = vbChecked Then
        strFaxed = "1"
    Else
        strFaxed = "0"
    End If
                    
    If chkwentalk.Value = vbChecked Then
        strwentalk = "1"
    Else
        strwentalk = "0"
    End If
                    
    If chkKTP.Value = vbChecked Then
        If StatusChekcBox = "" Then
            StatusChekcBox = "KTP "
        Else
            StatusChekcBox = StatusChekcBox + ",KTP "
        End If
        strKTP = "1"
    Else
        strKTP = "0"
    End If
                                       
    If chkpp.Value = vbChecked Then
        If StatusChekcBox = "" Then
            StatusChekcBox = "Surper "
        Else
            StatusChekcBox = StatusChekcBox + ",Surper "
        End If
        strSup = "1"
    Else
        strSup = "0"
    End If
                    
    If chkbillings.Value = vbChecked Then
        If StatusChekcBox = "" Then
            StatusChekcBox = "Billing "
        Else
            StatusChekcBox = StatusChekcBox + ",Billing"
        End If
        strBilling = "1"
    Else
        strBilling = "0"
    End If
                    
    If Check1.Value = vbChecked Then
        If StatusChekcBox = "" Then
            StatusChekcBox = "Other "
        Else
            StatusChekcBox = StatusChekcBox + ",Other"
        End If
        strOthers = "1"
    Else
        strOthers = "0"
    End If
    '====================== Cek Checkbox ======================================
                    
    '@@02-07-2012
    'tenor tidak boleh lebih dari 999
    If txttenor.Value > 999 Then
        MsgBox "Tenor maximal adalah 999! Send PTP gagal!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    Occupation() = Split(CmbOccupation.text, "-")
    Reason() = Split(CmbReason.text, "-")
    
    CMDSQL = "insert into tblsendptp (jenis_ptp,date_payment_effective,"
    CMDSQL = CMDSQL + "total_amount_deal,tenor,pembayaran_via,tgl_tagih,"
    CMDSQL = CMDSQL + "custid,agent,balance,principal,pembayaran_awal,"
    CMDSQL = CMDSQL + "nttlpayment,ndownpay,nfuturepay,ncharge,ndiscountamt,"
    CMDSQL = CMDSQL + "vosbalance,vosprincipal,vjust,nbalance,nperiod,"
    CMDSQL = CMDSQL + "vcustname,chkfaxed,nprincipal,chkwentalking,chkktp,"
    CMDSQL = CMDSQL + "chksup,chkbillings,chkothers,payment_after_tenor,dob"
    CMDSQL = CMDSQL + ",ket_other,payment_handle,"
    
    CMDSQL = CMDSQL + "occupation,reason"
    CMDSQL = CMDSQL + ",cek_pernyataan"
    
    CMDSQL = CMDSQL + ") values ('"
    CMDSQL = CMDSQL + Trim(CmbJenisPTP.text) + "','"
    CMDSQL = CMDSQL + Format(TxtPeymentEffective.Value, "yyyy-mm-dd") + "','"
    CMDSQL = CMDSQL + CStr(txtPayment.Value) + "','"
    CMDSQL = CMDSQL + CStr(txttenor.Value) + "','"
    CMDSQL = CMDSQL + CmbViaPtp.text + "','"
    CMDSQL = CMDSQL + Format(TxtTglTagih.Value, "yyyy-mm-dd") + "','"
    CMDSQL = CMDSQL + FrmCC_Colection.lblCustId.Caption + "','"
    'CMDSQL = CMDSQL + MDIForm1.Text1.Text + "','"
    CMDSQL = CMDSQL + FrmCC_Colection.lblaoc.Caption + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(FrmCC_Colection.TDB_cur_bal.Value), "0", FrmCC_Colection.TDB_cur_bal.Value)) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(FrmCC_Colection.TxtCurpri.Value), "0", FrmCC_Colection.TxtCurpri.Value)) + "','"
    CMDSQL = CMDSQL + CStr(txtPembayaranAwal.Value) + "','"
    CMDSQL = CMDSQL + CStr(lblLastPay.Value) + "','"
    CMDSQL = CMDSQL + CStr(txtdownpayment.Value) + "','"
    'CMDSQL = CMDSQL + "0','"
    CMDSQL = CMDSQL + CStr(txtfuture.Value) + "','"
    CMDSQL = CMDSQL + CStr(txtcharge.Value) + "','"
    CMDSQL = CMDSQL + CStr(txtdiscount.Value) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(txtfrombalancepersen.text), "", txtfrombalancepersen.text) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(txtpersenprincipal.text), "", txtpersenprincipal.text) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(TxtJustification.text), "", TxtJustification.text) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(txtbalance.Value), "", CStr(txtbalance.Value)) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(tdbisnstallment.Value), "", CStr(tdbisnstallment.Value)) + "','"
    CMDSQL = CMDSQL + FrmCC_Colection.lblNama.Caption + "','"
    CMDSQL = CMDSQL + strFaxed + "','"
    CMDSQL = CMDSQL + IIf(IsNull(txtprincipal.Value), "", CStr(txtprincipal.Value)) + "','"
    CMDSQL = CMDSQL + strwentalk + "','"
    CMDSQL = CMDSQL + strKTP + "','"
    CMDSQL = CMDSQL + strSup + "','"
    CMDSQL = CMDSQL + strBilling + "','"
    CMDSQL = CMDSQL + strOthers + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(TxtPayAfterTenor.Value), "0", TxtPayAfterTenor.Value)) + "',"
    CMDSQL = CMDSQL + IIf(TxtDob.text = "", "null", "'" + TxtDob.text + "'") + ",'"
    CMDSQL = CMDSQL + IIf(IsNull(txtothers.text), "", txtothers.text) + "','"
    CMDSQL = CMDSQL + CStr(Trim(CmbPaymentHandle.text)) + "','"
    CMDSQL = CMDSQL + Occupation(0) + "','"
    CMDSQL = CMDSQL + Reason(0) + "','"
    CMDSQL = CMDSQL + CStr(ChkPersetujuan.Value) + "')"
    M_OBJCONN.execute CMDSQL
    
    CMDSQL = "select * from usertbl where userid='"
    CMDSQL = CMDSQL + MDIForm1.Text1.text + "'"
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_OBJRS.RecordCount > 0 Then
        Dim M_Objrs_Pesan As ADODB.Recordset
        
        With FrmCC_Colection
            
            Remarks = "Agent anda: " & .lblaoc.Caption & vbCrLf
            Remarks = Remarks & "Mengirimkan permintaan account PTP untuk: " & vbCrLf
            Remarks = Remarks & "Jenis PTP :" & UCase(CmbJenisPTP.text) & vbCrLf
            Remarks = Remarks & "Custid : " & CStr(.lblCustId.Caption) & vbCrLf
            Remarks = Remarks & "Nama CH : " & CStr(.lblNama.Caption) & vbCrLf
            Remarks = Remarks & "Tgl.Effective :" & CStr(Format(TxtPeymentEffective.Value, "yyyy-mm-dd")) & vbCrLf
            Remarks = Remarks & "Amount Deal :" & txtPayment.Value & vbCrLf
            Remarks = Remarks & "Tenor :" & CStr(txttenor.Value) & vbCrLf
            Remarks = Remarks & "Via :" & CmbViaPtp.text & vbCrLf
            Remarks = Remarks & "Tgl.Tagih :" & CStr(Format(TxtTglTagih.Value, "yyyy-mm-dd")) & vbCrLf
            
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL & "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL & M_OBJRS("team") & "','"
            CMDSQL = CMDSQL & Format(Now(), "yyyymmdd") & "','"
            CMDSQL = CMDSQL & MDIForm1.Text1.text & "','"
            CMDSQL = CMDSQL & CStr(MDIForm1.Winsock1.LocalIP) & "','"
            CMDSQL = CMDSQL & Remarks & "')"
            M_OBJCONN.execute CMDSQL
            
             '15 Juni 2012, SPV tidak usah dikasih Pesan
'            'Buat Kirim Pesan Ke SPV
'            CMDSQL = "select userid from usertbl where usertype in ('11','20','25') and userid is not null "
'            Set M_Objrs_Pesan = New ADODB.Recordset
'            M_Objrs_Pesan.CursorLocation = adUseClient
'            M_Objrs_Pesan.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If M_Objrs_Pesan.RecordCount > 0 Then
'                While Not M_Objrs_Pesan.EOF
'                    CMDSQL = "insert into msgtbl "
'                    CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
'                    CMDSQL = CMDSQL + M_Objrs_Pesan("userid") + "','"
'                    CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
'                    CMDSQL = CMDSQL + MDIForm1.Text1.Text + "','"
'                    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
'                    CMDSQL = CMDSQL + Remarks + "')"
'                    M_OBJCONN.Execute CMDSQL
'                    M_Objrs_Pesan.MoveNext
'                Wend
'            End If
'            Set M_Objrs_Pesan = Nothing
            
        End With
    End If
    Set M_OBJRS = Nothing
    
    MsgBox "Permintaan PTP anda berhasil disimpan dan dikirimkan ke TL anda!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

Private Sub CmdUnCekAll_Click()
    Dim W As Integer
    
    If LvLogPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!"
        Exit Sub
    End If
    
    For W = 1 To LvLogPTP.ListItems.Count
        LvLogPTP.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    'CmbJenisPTP.Text = "PTP 1X BAYAR"
    CmbJenisPTP.text = "PTP No Discount"
    Call HeaderLog
    Call isilog
    
    Label5.text = IIf(FrmCC_Colection.TDB_cur_bal.ValueIsNull, "0", FrmCC_Colection.TDB_cur_bal.Value)
    Label8.text = IIf(FrmCC_Colection.TxtCurpri.ValueIsNull, "0", FrmCC_Colection.TxtCurpri.Value)
    txtbalance.Value = IIf(FrmCC_Colection.TDB_cur_bal.ValueIsNull, "0", FrmCC_Colection.TDB_cur_bal.Value)
    txtprincipal.Value = IIf(FrmCC_Colection.TxtCurpri.ValueIsNull, "0", FrmCC_Colection.TxtCurpri.Value)
    TxtDob.text = IIf(IsNull(FrmCC_Colection.LblDOB.Caption), "", Format(FrmCC_Colection.LblDOB.Caption, "yyyy-mm-dd"))
    
    '@@ 02072012, Jika yang login Admin/SPV Balance dapat diedit
    If UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Then
       txtbalance.Enabled = True
    ElseIf UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Or _
           UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        txtbalance.Enabled = False
    End If
End Sub

Private Sub CariTanggalTagih()
    Dim CMDSQL As String
    Dim M_OBJRS As ADODB.Recordset
    Dim TglPaymentEffective As String
    
    If IsNull(TxtPeymentEffective.Value) = True Then
        MsgBox "Payment effective tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    TglPaymentEffective = Format(TxtPeymentEffective.Value, "yyyy-mm-dd")
    
    CMDSQL = "Select  date('" + TglPaymentEffective + "')-"
    If UCase(Trim(CmbViaPtp.text)) = "HSBC" Then
        CMDSQL = CMDSQL + "1"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "BERSAMA" Then
        CMDSQL = CMDSQL + "1"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "KANTOR POS" Then
        CMDSQL = CMDSQL + "3"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "PUM" Then
        CMDSQL = CMDSQL + "1"
    Else
        CMDSQL = CMDSQL + "3"
    End If
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    On Error GoTo salah
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtTglTagih.Value = Format(M_OBJRS(0), "mm/dd/yyyy")
    
    Set M_OBJRS = Nothing
    Exit Sub
salah:
    MsgBox "Ada Error: " & err.Description
End Sub







Private Sub txtothers_Change()
    LblJumlah.Caption = Len(txtothers.text)
End Sub

Private Sub TxtPayAfterTenor_Change()
    Call CariTenor
    Call PaymentAfterTenor
End Sub

Private Sub txtPayment_Change()
    lblLastPay.Value = txtPayment.Value
    txtPembayaranAwal.Value = txtPayment.Value
    txtdownpayment.Value = txtPembayaranAwal.Value
    
    Call CariTenor
    Call PaymentAfterTenor
    
    '@@22062012 Jika payment=pembayaran awal set tenor=1
    If txtPayment.Value = txtPembayaranAwal Then
        txttenor.Value = 1
    End If
    
    '@@22062012 Jika Payment=Balance di Database maka otomatis jadi PTP No Discount
    If txtPayment.Value = txtbalance.Value Then
        CmbJenisPTP.text = "PTP No Discount"
    End If
    
    If txtPayment < txtbalance.Value Then
        CmbJenisPTP.text = "PTP Discount"
    End If
    
    If txtPayment.Value > txtbalance.Value Then
        MsgBox "Total Amount Deal tidak boleh lebih besar dari balance!", vbOKOnly + vbInformation, "Informasi"
        txtPayment.Value = 1
        Exit Sub
    End If
    
    If txtPembayaranAwal.Value = txtPayment.Value Then
        txtdownpayment.Value = 0
    Else
        txtdownpayment.Value = txtPembayaranAwal.Value
    End If
End Sub

Private Sub txtPembayaranAwal_Change()
    If txtPembayaranAwal.Value > txtPayment.Value Then
        MsgBox "Pembayaran Awal tidak boleh lebih besar dari total payment effective!"
        txtPembayaranAwal.Value = 0
        txtdownpayment.Value = 0
        Exit Sub
    End If
    
    If txtPembayaranAwal.Value = txtPayment.Value Then
        txtdownpayment.Value = 0
    Else
        txtdownpayment.Value = txtPembayaranAwal.Value
    End If
    
    Call CariTenor
    Call PaymentAfterTenor
    
    
    If txtPayment.Value = txtPembayaranAwal Then
        txttenor.Value = 1
    End If
End Sub

Private Sub TxtPeymentEffective_Change()
      Call CariTanggalTagih
End Sub

Private Sub TxtPeymentEffective_Click()
    Call CariTanggalTagih
End Sub


Private Sub HeaderLog()
    With LvLogPTP.ColumnHeaders
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
    End With
End Sub

Private Sub isilog()
    Dim CMDSQL As String
    Dim M_OBJRS As ADODB.Recordset
    Dim listitem As listitem
    
    CMDSQL = "select * from tblsendptp where agent='"
    CMDSQL = CMDSQL + MDIForm1.Text1.text + "' order by tgldata desc"
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvLogPTP.ListItems.clear
    TxtJumlah.text = M_OBJRS.RecordCount
    If M_OBJRS.RecordCount > 0 Then
        Dim STATUS As String
        While Not M_OBJRS.EOF
            Set listitem = LvLogPTP.ListItems.ADD(, , M_OBJRS("id"))
                listitem.SubItems(1) = IIf(IsNull(M_OBJRS("jenis_ptp")), "", M_OBJRS("jenis_ptp"))
                listitem.SubItems(2) = IIf(IsNull(M_OBJRS("custid")), "", M_OBJRS("custid"))
                listitem.SubItems(3) = IIf(IsNull(M_OBJRS("vcustname")), "", M_OBJRS("vcustname"))
                
                If M_OBJRS("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_OBJRS("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_OBJRS("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                listitem.SubItems(4) = STATUS
                listitem.SubItems(5) = IIf(IsNull(M_OBJRS("tgl_approve")), "", Format(M_OBJRS("tgl_approve"), "yyyy-mm-dd"))
                listitem.SubItems(6) = IIf(IsNull(M_OBJRS("date_payment_effective")), "", Format(M_OBJRS("date_payment_effective"), "yyyy-mm-dd"))
                listitem.SubItems(7) = IIf(IsNull(M_OBJRS("total_amount_deal")), "", Format(M_OBJRS("total_amount_deal"), "##,###"))
                listitem.SubItems(8) = IIf(IsNull(M_OBJRS("tenor")), "", Format(M_OBJRS("tenor"), "##,###"))
                listitem.SubItems(9) = IIf(IsNull(M_OBJRS("pembayaran_via")), "", M_OBJRS("pembayaran_via"))
                listitem.SubItems(10) = IIf(IsNull(M_OBJRS("tgl_tagih")), "", Format(M_OBJRS("tgl_tagih"), "yyyy-mm-dd"))
                listitem.SubItems(11) = IIf(IsNull(M_OBJRS("principal")), "", Format(M_OBJRS("principal"), "##,###"))
                listitem.SubItems(12) = IIf(IsNull(M_OBJRS("balance")), "", Format(M_OBJRS("balance"), "##,###"))
                listitem.SubItems(13) = IIf(IsNull(M_OBJRS("pembayaran_awal")), "", Format(M_OBJRS("pembayaran_awal"), "##,###"))
                listitem.SubItems(14) = IIf(IsNull(M_OBJRS("principal")), "", Format(M_OBJRS("principal"), "##,###"))
                listitem.SubItems(15) = IIf(IsNull(M_OBJRS("nttlpayment")), "", Format(M_OBJRS("nttlpayment"), "##,###"))
                listitem.SubItems(16) = IIf(IsNull(M_OBJRS("ndownpay")), "", Format(M_OBJRS("ndownpay"), "##,###"))
                listitem.SubItems(17) = IIf(IsNull(M_OBJRS("ncharge")), "", Format(M_OBJRS("ncharge"), "##,###"))
                listitem.SubItems(18) = IIf(IsNull(M_OBJRS("ndiscountamt")), "", Format(M_OBJRS("ndiscountamt"), "##,###"))
                listitem.SubItems(19) = IIf(IsNull(M_OBJRS("vosbalance")), "", M_OBJRS("vosbalance"))
                listitem.SubItems(20) = IIf(IsNull(M_OBJRS("vosprincipal")), "", M_OBJRS("vosprincipal"))
                listitem.SubItems(21) = IIf(IsNull(M_OBJRS("vjust")), "", M_OBJRS("vjust"))
                listitem.SubItems(22) = IIf(IsNull(M_OBJRS("chkfaxed")), "", M_OBJRS("chkfaxed"))
                listitem.SubItems(23) = IIf(IsNull(M_OBJRS("chkwentalking")), "", M_OBJRS("chkwentalking"))
                listitem.SubItems(24) = IIf(IsNull(M_OBJRS("chkktp")), "", M_OBJRS("chkktp"))
                listitem.SubItems(25) = IIf(IsNull(M_OBJRS("chksup")), "", M_OBJRS("chksup"))
                listitem.SubItems(26) = IIf(IsNull(M_OBJRS("chkbillings")), "", M_OBJRS("chkbillings"))
                listitem.SubItems(27) = IIf(IsNull(M_OBJRS("chkothers")), "", M_OBJRS("chkothers"))
            M_OBJRS.MoveNext
        Wend
    End If
    Set M_OBJRS = Nothing
End Sub
Private Sub chkfaxed_Click()
    '@@ 08092011
    If chkfaxed.Value = vbChecked Or chkwentalk.Value = vbChecked Then
        chkKTP.Enabled = True
        chkpp.Enabled = True
        chkbillings.Enabled = True
        Check1.Enabled = True
    End If
    If chkfaxed.Value = vbUnchecked And chkwentalk.Value = vbUnchecked Then
        chkKTP.Enabled = False
        chkpp.Enabled = False
        chkbillings.Enabled = False
        Check1.Enabled = False
        
        chkKTP.Value = vbUnchecked
        chkpp.Value = vbUnchecked
        chkbillings.Value = vbUnchecked
        Check1.Value = vbUnchecked
    End If
End Sub


Private Sub chkwentalk_Click()
    '@@ 08092011
    '@@231012 Dinonaktifkan dulu
'    If chkfaxed.Value = vbChecked Or chkwentalk.Value = vbChecked Then
'        chkKTP.Enabled = True
'        chkpp.Enabled = True
'        chkbillings.Enabled = True
'        Check1.Enabled = True
'    End If
'    If chkfaxed.Value = vbUnchecked And chkwentalk.Value = vbUnchecked Then
'        chkKTP.Enabled = False
'        chkpp.Enabled = False
'        chkbillings.Enabled = False
'        Check1.Enabled = False
'
'        chkKTP.Value = vbUnchecked
'        chkpp.Value = vbUnchecked
'        chkbillings.Value = vbUnchecked
'        Check1.Value = vbUnchecked
'    End If
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        txtothers.Enabled = True
        txtothers.BackColor = vbWhite
    Else
        txtothers.Enabled = False
        txtothers.BackColor = &HC0C0C0
    End If
End Sub

'------------------------ BUAT CPA Perhitungan ----------------------------------------------
Private Sub txtbalance_Change()
    txtcharge.Value = txtbalance.Value - txtprincipal.Value
    txtdiscount.Value = txtbalance.Value - lblLastPay.Value
    If txtbalance.Value <> 0 Then
         txtfrombalancepersen.text = Round(((lblLastPay.Value / txtbalance.Value) - 1) * 100, 2)
    End If
    
    '@@ 12Juni2012, Jika Balance=0 maka persentase balance =0
    If txtbalance.Value = 0 Then
        txtfrombalancepersen.text = 0
    End If
    
    'Call PaymentAfterTenor
    
End Sub


Private Sub txtdownpayment_Change()
    txtfuture.Value = lblLastPay.Value - txtdownpayment.Value
    'Call PaymentAfterTenor
End Sub

Private Sub txtprincipal_Change()
    txtcharge.Value = txtbalance.Value - txtprincipal.Value
    If txtprincipal.Value <> 0 Then
        txtpersenprincipal.text = Round(((lblLastPay.Value / txtprincipal.Value) - 1) * 100, 2)
    End If
    
    
    '@@12 Juni 2012 Jika principal=0 maka persentase principal =0
    If txtprincipal.Value = 0 Then
        txtpersenprincipal.text = "0"
    End If
End Sub

Private Sub lblLastPay_Change()
    txtdiscount.Value = txtbalance.Value - lblLastPay.Value
    If txtbalance.Value <> 0 Then
        txtfrombalancepersen.text = Round(((lblLastPay.Value / txtbalance.Value) - 1) * 100, 2)
    End If
    If txtprincipal.Value <> 0 Then
        txtpersenprincipal.text = Round(((lblLastPay.Value / txtprincipal.Value) - 1) * 100, 2)
    End If
    
    
    '@@ 12Juni2012, Jika Balance=0 maka persentase balance=0. Jika Principal=0 maka persentase principal=0
    If txtbalance.Value = 0 Then
        txtfrombalancepersen.text = "0"
    End If
    If txtprincipal.Value = 0 Then
        txtpersenprincipal.text = "0"
    End If
    
    'Call PaymentAfterTenor
End Sub
'------------------------ Akhir BUAT CPA Perhitungan ----------------------------------------------

Private Sub txttenor_Change()
    On Error Resume Next
    tdbisnstallment.Value = txttenor.Value
    
    Call PaymentAfterTenor
    
    If txtPayment.Value = txtPembayaranAwal Then
        txttenor.Value = 1
    End If
End Sub

Private Sub PaymentAfterTenor()
    Dim PayAfterTenor As Double
    
    PayAfterTenor = 0
    If (tdbisnstallment - 1) = 0 Then
        PayAfterTenor = 0
    Else
        PayAfterTenor = (lblLastPay.Value - txtdownpayment.Value) / (tdbisnstallment - 1)
    End If
    On Error Resume Next
    'TxtPayAfterTenor.Value = PayAfterTenor
    TxtPaymentMonthSebenarnya.Value = Ceiling(PayAfterTenor)
End Sub

Private Sub CariTenor()
    Dim Payment As Double
    Dim DownPayment As Double
    Dim Tenor As Double
    Dim PaymentAfterTenor As Double
    
    Payment = txtPayment.Value
    DownPayment = txtdownpayment.Value
    PaymentAfterTenor = TxtPayAfterTenor.Value
    
    On Error Resume Next
    Tenor = ((Payment - DownPayment) / PaymentAfterTenor) + 1
    txttenor.Value = Ceiling(Tenor)
End Sub

Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function
