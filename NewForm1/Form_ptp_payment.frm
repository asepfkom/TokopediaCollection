VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form Form_ptp_payment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form PTP And Payment"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   19935
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   19935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   5640
      TabIndex        =   35
      Top             =   7920
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmd_change 
         BackColor       =   &H0000FF00&
         Caption         =   "Change Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox txt_amount 
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
         Left            =   0
         TabIndex        =   37
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmd_amount 
         BackColor       =   &H0000FF00&
         Caption         =   "Change Amount"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   480
         Width           =   1695
      End
      Begin TDBDate6Ctl.TDBDate txt_date 
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "Form_ptp_payment.frx":0000
         Caption         =   "Form_ptp_payment.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_ptp_payment.frx":0184
         Keys            =   "Form_ptp_payment.frx":01A2
         Spin            =   "Form_ptp_payment.frx":0200
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
   End
   Begin VB.TextBox txt_total_payment 
      Height          =   285
      Left            =   3960
      TabIndex        =   34
      Text            =   "0"
      Top             =   8880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt_total_ptp 
      Height          =   285
      Left            =   3840
      TabIndex        =   33
      Text            =   "0"
      Top             =   8880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmd_download_payment 
      BackColor       =   &H0000FF00&
      Caption         =   "Download Payment To Excel"
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton cmd_download 
      BackColor       =   &H0000FF00&
      Caption         =   "Download PTP To Excel"
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd_count 
      BackColor       =   &H0000FF00&
      Caption         =   "Count . . ."
      Enabled         =   0   'False
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
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chk_team 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3840
      TabIndex        =   16
      Top             =   1440
      Width           =   195
   End
   Begin VB.ComboBox cmb_team 
      Enabled         =   0   'False
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
      ItemData        =   "Form_ptp_payment.frx":0228
      Left            =   1200
      List            =   "Form_ptp_payment.frx":022A
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmd_paid 
      BackColor       =   &H0000FF00&
      Caption         =   "Paid"
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
      TabIndex        =   11
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monthly Payment"
      Height          =   5535
      Left            =   9480
      TabIndex        =   9
      Top             =   2280
      Width           =   10455
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
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_payment 
         BackColor       =   &H0000FF00&
         Caption         =   "Show Payment"
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
         TabIndex        =   12
         Top             =   5040
         Width           =   1575
      End
      Begin MSComctlLib.ListView LvPayment 
         Height          =   4350
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   7673
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
      Begin VB.Label lblpayment 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL : IDR"
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
         Left            =   3960
         TabIndex        =   13
         Top             =   5160
         Width           =   4455
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Close"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmd_showptp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show PTP"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cmb_agent 
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
      ItemData        =   "Form_ptp_payment.frx":022C
      Left            =   1200
      List            =   "Form_ptp_payment.frx":022E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monthly PTP"
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   9255
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
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   4350
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
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
      Begin VB.Label lbltotal 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL : IDR"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   5040
         Width           =   4815
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
         Top             =   5040
         Width           =   2655
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
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
      CustomFormat    =   "MMMM-yyyy"
      Format          =   122355715
      CurrentDate     =   41610
   End
   Begin TDBDate6Ctl.TDBDate tgl_mulai1 
      Height          =   375
      Left            =   6000
      TabIndex        =   25
      Top             =   1320
      Width           =   1845
      _Version        =   65536
      _ExtentX        =   3254
      _ExtentY        =   661
      Calendar        =   "Form_ptp_payment.frx":0230
      Caption         =   "Form_ptp_payment.frx":0348
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_ptp_payment.frx":03B4
      Keys            =   "Form_ptp_payment.frx":03D2
      Spin            =   "Form_ptp_payment.frx":0430
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
      Left            =   8400
      TabIndex        =   26
      Top             =   1320
      Width           =   1845
      _Version        =   65536
      _ExtentX        =   3254
      _ExtentY        =   661
      Calendar        =   "Form_ptp_payment.frx":0458
      Caption         =   "Form_ptp_payment.frx":0570
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_ptp_payment.frx":05DC
      Keys            =   "Form_ptp_payment.frx":05FA
      Spin            =   "Form_ptp_payment.frx":0658
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   6000
      TabIndex        =   30
      Top             =   1800
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lbl_total_keseluruhan 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIDENT : IDR"
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
      Left            =   10440
      TabIndex        =   32
      Top             =   8520
      Width           =   4815
   End
   Begin VB.Label Label2 
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
      Left            =   7920
      TabIndex        =   27
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lbl_total_hitung_payment 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL : IDR"
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
      Left            =   10440
      TabIndex        =   21
      Top             =   8160
      Width           =   4815
   End
   Begin VB.Label lbl_jumlah_data_dipilih_payment 
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
      Left            =   10440
      TabIndex        =   20
      Top             =   7800
      Width           =   2655
   End
   Begin VB.Label lbl_jumlah_data_dipilih 
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
      Left            =   1560
      TabIndex        =   19
      Top             =   8280
      Width           =   2655
   End
   Begin VB.Label lbl_total_hitung 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL : IDR"
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
      Left            =   1560
      TabIndex        =   18
      Top             =   7920
      Width           =   4815
   End
   Begin VB.Label lbl_team 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Team  :"
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
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Agent :"
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
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "PTP - PAYMENT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   0
      Top             =   60
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   120
      Picture         =   "Form_ptp_payment.frx":0680
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "Form_ptp_payment.frx":118A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20400
   End
End
Attribute VB_Name = "Form_ptp_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_list As ADODB.Recordset
Dim f_team As Boolean

Private Sub koneksi()
    Set Rs_list = New ADODB.Recordset
    Rs_list.CursorLocation = adUseClient
    Rs_list.ActiveConnection = M_OBJCONN
    Rs_list.CursorType = adOpenDynamic
    Rs_list.LockType = adLockOptimistic
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        chk_team.Value = 0
    End If
End Sub

Private Sub cek_all_payment_Click()
    Dim r As Integer
        
    If cek_all_payment.Value = vbChecked Then
        If LvPayment.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To LvPayment.ListItems.Count
            LvPayment.ListItems(r).Checked = True
        Next r
        Call Hitung_Payment_Dipilih
    Else
        For r = 1 To LvPayment.ListItems.Count
            LvPayment.ListItems(r).Checked = False
        Next r
        Call Hitung_Payment_Dipilih
    End If
    
End Sub

Private Sub cek_all_ptp_Click()
    Dim r As Integer
        
    If cek_all_ptp.Value = vbChecked Then
        If LvPTP.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To LvPTP.ListItems.Count
            LvPTP.ListItems(r).Checked = True
        Next r
        Call cmd_count_Click
    Else
        For r = 1 To LvPTP.ListItems.Count
            LvPTP.ListItems(r).Checked = False
        Next r
        Call cmd_count_Click
    End If
End Sub

Private Sub chk_team_Click()
    If chk_team.Value = vbChecked Then
        Call Isi_TL
        cmb_agent.ListIndex = 0
        cmb_agent.Enabled = False
        cmb_team.Enabled = True
        f_team = True
        Check1.Value = 0
    Else
        cmb_agent.Enabled = True
        cmb_team.Enabled = False
        cmb_team.ListIndex = 0
        f_team = False
    End If
End Sub



Private Sub Isi_TL()
    If Rs_list.state = 1 Then Rs_list.Close
    
    'If Left(MDIForm1.Text2.text, 2) = "AM" Then
    '    Rs_list.Open "SELECT DISTINCT team FROM usertbl where team ilike  'TL%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')"
    'Else
        Rs_list.Open "SELECT DISTINCT team FROM usertbl where team ilike  'TL%' "
    'End If
    cmb_team.AddItem " "
    
    While Not Rs_list.EOF
        cmb_team.AddItem Rs_list("team")
        Rs_list.MoveNext
    Wend
End Sub

Private Sub cmd_amount_Click()
    Dim w As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim cmdsql As String
    Dim tanggal_bayar As Date
    Dim id_ptp1 As String
    Dim CustId As String
    
    If txt_amount.text = "" Then
        MsgBox "Masukkan Jumlah Amount Yang Baru", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    
    a = MsgBox("Apakah Anda Yakin Akan Merubah Amount PTP?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses Dibatalkan!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek > 1 Then
        MsgBox "Anda Tidak Boleh Memilih Lebih Dari 1 PTP!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    If cek = 0 Then
        MsgBox "Check PTP Terlebih Dahulu!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    

    id_ptp1 = LvPTP.SelectedItem.SubItems(4)
    CustId = LvPTP.SelectedItem.text
    
    For w = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(w).Checked = True Then
            cmdsql = "UPDATE tblnegoptp SET promisepay = '" & txt_amount.text & "' "
            cmdsql = cmdsql + " WHERE id = '" & id_ptp1 & "' "
            M_OBJCONN.Execute cmdsql
            
            cmdsql = "UPDATE mgm SET amountptp = '" & txt_amount.text & "' "
            cmdsql = cmdsql + " WHERE custid = '" & CustId & "' "
            M_OBJCONN.Execute cmdsql
        End If
    Next w
    
    MsgBox "Amount PTP Berhasil Di-Ubah!", vbOKOnly + vbInformation, "Informasi"
    LvPTP.ListItems(1).Checked = False
    Call IsiAccountPTP
End Sub

Private Sub cmd_change_Click()
    Dim w As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim cmdsql As String
    Dim tanggal_bayar As Date
    Dim id_ptp1 As String
    Dim CustId As String
    
    If txt_date.ValueIsNull Then
        MsgBox "Masukkan Tanggal PTP Yang Baru", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    
    a = MsgBox("Apakah Anda Yakin Akan Merubah Tanggal PTP?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses Dibatalkan!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
'    If cek > 1 Then
'        MsgBox "Anda Tidak Boleh Memilih Lebih Dari 1 PTP!", vbOKOnly + vbExclamation, "Peringatan"
'        Exit Sub
'    End If
    
    If cek = 0 Then
        MsgBox "Check PTP Terlebih Dahulu!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    

    id_ptp1 = LvPTP.SelectedItem.SubItems(4)
    CustId = LvPTP.SelectedItem.text
    
    For w = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(w).Checked = True Then
            cmdsql = "UPDATE tblnegoptp SET promisedate = '" & txt_date.text & "' "
            cmdsql = cmdsql + " WHERE id = '" & LvPTP.ListItems(w).SubItems(4) & "' "
            M_OBJCONN.Execute cmdsql
            
            cmdsql = "UPDATE mgm SET dateptp = '" & txt_date.text & "' "
            cmdsql = cmdsql + " WHERE custid = '" & LvPTP.ListItems(w).text & "' "
            M_OBJCONN.Execute cmdsql
        End If
    Next w
    
    MsgBox "Tanggal PTP Berhasil Di-Ubah!", vbOKOnly + vbInformation, "Informasi"
    LvPTP.ListItems(1).Checked = False
    Call IsiAccountPTP
End Sub

Private Sub cmd_count_Click()
On Error GoTo bawah
    Dim w As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim janji_bayar As Double
    Dim total_janji_bayar As Double
    
    cek = 0
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
      

    
    For w = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(w).Checked = True Then
'            janji_bayar = LvPTP.SelectedItem.SubItems(2)
            janji_bayar = LvPTP.ListItems(w).SubItems(2)
            
            total_janji_bayar = total_janji_bayar + janji_bayar
        End If
    Next w
    
    lbl_total_hitung.Caption = "TOTAL : IDR " + Format(total_janji_bayar, "##,###")
    lbl_jumlah_data_dipilih.Caption = "Jumlah " & cek & " Rows"
bawah:
End Sub

Private Sub cmd_download_Click()
    Call My_Export_Excel_PTP
End Sub

Private Sub My_Export_Excel_PTP()
    Dim a           As Long
    Dim B           As Long
    Dim ExlObj      As Excel.Application
    Dim listcustid  As String
    Dim rs          As ADODB.Recordset
    Dim RS2         As ADODB.Recordset
    Dim iRow        As Integer
    Dim i           As Integer
    Dim sQuery      As String
    Dim totalcall   As Double
    Dim totaldata   As Double
    Dim ratarata   As Double
    Dim agent As String
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim id_payment As String
        
    jam_mulai = "00:00:00"
    jam_selesai = "23:59:59"
    
    tgl_mulai = Format(tgl_mulai1.Value, "YYYY-MM-DD")
    tgl_akhir = Format(tgl_akhir1.Value, "YYYY-MM-DD")
'    tgl_mulai = tgl_mulai1.Value
'    tgl_akhir = tgl_akhir1.Value
    
    agent = cmb_agent.text
    
    If cmb_agent = " " Then
        agent = ""
    End If
    
    If cmb_agent = "ALL" Then
'        If Left(MDIForm1.Text2.text, 2) = "AM" Then
'            sQuery = "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( "
'            sQuery = sQuery + " SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) "
'            sQuery = sQuery + " AND payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
'      '      sQuery = sQuery + " AND payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
'            sQuery = sQuery + " order by custid) As a RIGHT JOIN ( "
'            sQuery = sQuery + " SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp "
'            sQuery = sQuery + " WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) "
'            sQuery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & "' "
'            'squery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' "
'            sQuery = sQuery + " AND '" & tgl_akhir & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
'            Set rs = New ADODB.Recordset
'            rs.CursorLocation = adUseClient
'            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        Else
            sQuery = "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name, acc_type FROM( "
            sQuery = sQuery + " SELECT Distinct custid FROM tbllunas WHERE agent ilike 'D%' and char_length(agent) = 4 "
            'sQuery = sQuery + " AND payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + " AND payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
            sQuery = sQuery + " order by custid) As a RIGHT JOIN ( "
            sQuery = sQuery + " SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp "
            sQuery = sQuery + " WHERE mgm.custid = tblnegoptp.custid AND agent ilike 'D%' and char_length(agent) = 4 "
            sQuery = sQuery + " AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' "
    '        sQuery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' "
            sQuery = sQuery + " AND '" & tgl_akhir & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        End If
    Else
        If Check1.Value = 1 Then
            sQuery = "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( "
            sQuery = sQuery + " SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) "
            sQuery = sQuery + " AND payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
      '      sQuery = sQuery + " AND payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + " order by custid) As a RIGHT JOIN ( "
            sQuery = sQuery + " SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp "
            sQuery = sQuery + " WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) "
            sQuery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & "' "
            'squery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' "
            sQuery = sQuery + " AND '" & tgl_akhir & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        ElseIf f_team = False Then
            sQuery = "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( "
            sQuery = sQuery + " SELECT Distinct custid FROM tbllunas WHERE agent = '" & agent & "' "
            sQuery = sQuery + " AND payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
      '      sQuery = sQuery + " AND payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + " order by custid) As a RIGHT JOIN ( "
            sQuery = sQuery + " SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp "
            sQuery = sQuery + " WHERE mgm.custid = tblnegoptp.custid AND agent = '" & agent & "' "
            sQuery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & "' "
            'squery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' "
            sQuery = sQuery + " AND '" & tgl_akhir & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        ElseIf f_team = True Then
            sQuery = "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( "
            sQuery = sQuery + " SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') "
            sQuery = sQuery + " AND payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND  '" & tgl_akhir & "' "
'            sQuery = sQuery + " AND payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND  '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + " order by custid) As a RIGHT JOIN ( "
            sQuery = sQuery + " SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp "
            sQuery = sQuery + " WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') "
            sQuery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & "' "
            'sQuery = sQuery + " AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' "
            sQuery = sQuery + " AND '" & tgl_akhir & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        End If
    End If
   
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:N1").MergeCells = True
    'ExlObj.Range("A2:N2").MergeCells = True
    ExlObj.Range("A4:N4").Font.Bold = True
    
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "List PTP - Tanggal " & Format(tgl_mulai1.Value, "DD-MM-YYYY") & " Sampai " & Format(tgl_akhir1.Value, "DD-MM-YYYY")
        .Cells(1, 1).Font.Name = "Verdana"
        .Cells(1, 1).Font.Bold = True
        .Cells(4, 1).Value = "NO"
        .Cells(4, 2).Value = "CARD NUMBER"
        .Cells(4, 3).Value = "CH NAME"
        .Cells(4, 4).Value = "AGENT"
        .Cells(4, 5).Value = "PROMISEPAY"
        .Cells(4, 6).Value = "PROMISEDATE"
        .Cells(4, 7).Value = "PRODUCT"

        iRow = 4
        If rs.RecordCount > 0 Then
            ProgressBar1.Max = rs.RecordCount
            i = 0
            Do Until rs.EOF
                i = i + 1
                iRow = iRow + 1
                ProgressBar1.Value = rs.Bookmark
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 3).Value = IIf(IsNull(rs!Name), "", rs!Name)
                .Cells(iRow, 4).Value = IIf(IsNull(rs!agent), "", rs!agent)
                .Cells(iRow, 5).Value = Format(IIf(IsNull(rs!PromisePay), "", rs!PromisePay), "##,###")
                .Cells(iRow, 6).Value = IIf(IsNull(rs("PromiseDate")), "", Format(rs("PromiseDate"), "DD-MM-YYYY"))
                .Cells(iRow, 7).Value = IIf(IsNull(rs!acc_type), "", rs!acc_type)
                rs.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 14
            ExlObj.Cells(4, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
        ProgressBar1.Value = 0
        cmd_download.Enabled = True
    
        Set ExlObj = Nothing
        Set rs = Nothing

        'StartMeUp (Txtlocation.Text)
        'FILL COLOR CELL
        'ExlObj.Range(.Cells(NoUrut, 1), .Cells(NoUrut, 7)).Interior.Color = RGB(6, 207, 250)
    End With
End Sub

Private Sub cmd_download_payment_Click()
    Call My_Export_Excel_Payment
End Sub

Private Sub My_Export_Excel_Payment()
    Dim a           As Long
    Dim B           As Long
    Dim ExlObj      As Excel.Application
    Dim listcustid  As String
    Dim rs          As ADODB.Recordset
    Dim RS2         As ADODB.Recordset
    Dim iRow        As Integer
    Dim i           As Integer
    Dim sQuery      As String
    Dim totalcall   As Double
    Dim totaldata   As Double
    Dim ratarata   As Double
    Dim agent As String
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim id_payment As String
        
    jam_mulai = "00:00:00"
    jam_selesai = "23:59:59"
    
    tgl_mulai = Format(tgl_mulai1.Value, "YYYY-MM-DD")
    tgl_akhir = Format(tgl_akhir1.Value, "YYYY-MM-DD")
'    tgl_mulai = tgl_mulai1.Value
'    tgl_akhir = tgl_akhir1.Value
    
    agent = cmb_agent.text
    
    If cmb_agent = " " Then
        agent = ""
    End If
    
    If cmb_agent = "ALL" Then
'        If Left(MDIForm1.Text2.text, 2) = "AM" Then
'            sQuery = "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( "
'            sQuery = sQuery + "SELECT custid, paydate, payment, agent, id FROM tbllunas  "
'            sQuery = sQuery + "WHERE Payment > 100 AND date(paydate) between  '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
''            sQuery = sQuery + "WHERE Payment > 100 AND paydate between  '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
'            sQuery = sQuery + "AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as a left join  "
'            sQuery = sQuery + "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as b on a.custid = b.custid "
'            Set rs = New ADODB.Recordset
'            rs.CursorLocation = adUseClient
'            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        Else
            sQuery = "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( "
            sQuery = sQuery + " SELECT custid, paydate, payment, agent, id FROM tbllunas  "
            sQuery = sQuery + " WHERE Payment > 100 AND date(paydate) between  '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
    '        sQuery = sQuery + " WHERE Payment > 100 AND paydate between  '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + " AND agent ilike 'D%' and char_length(agent) = 4) as a left join  "
            sQuery = sQuery + " (SELECT custid, name, acc_type from mgm where agent ilike 'D%' and char_length(agent) = 4) as b on a.custid = b.custid "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        End If
    Else
        If Check1.Value = 1 Then
            sQuery = "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( "
            sQuery = sQuery + "SELECT custid, paydate, payment, agent, id FROM tbllunas  "
            sQuery = sQuery + "WHERE Payment > 100 AND date(paydate) between  '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
'            sQuery = sQuery + "WHERE Payment > 100 AND paydate between  '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + "AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as a left join  "
            sQuery = sQuery + "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as b on a.custid = b.custid "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        ElseIf f_team = False Then
            sQuery = "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( "
            sQuery = sQuery + "SELECT custid, paydate, payment, agent, id FROM tbllunas  "
            sQuery = sQuery + "WHERE Payment > 100 AND date(paydate) between  '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
'            sQuery = sQuery + "WHERE Payment > 100 AND paydate between  '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + "AND agent = '" & cmb_agent.text & "') as a left join  "
            sQuery = sQuery + "(SELECT custid, name, acc_type from mgm where agent = '" & cmb_agent.text & "') as b on a.custid = b.custid "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        ElseIf f_team = True Then
            sQuery = "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ("
            sQuery = sQuery + "SELECT custid, paydate, payment, agent, id FROM tbllunas  "
            sQuery = sQuery + "WHERE Payment > 100 AND date(paydate) between  '" & tgl_mulai & "' AND '" & tgl_akhir & "' "
'            sQuery = sQuery + "WHERE Payment > 100 AND paydate between  '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
            sQuery = sQuery + "AND agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%')) as a left join  "
            sQuery = sQuery + "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%')) as b on a.custid = b.custid "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        End If
    End If
   
    
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:N1").MergeCells = True
    'ExlObj.Range("A2:N2").MergeCells = True
    ExlObj.Range("A4:N4").Font.Bold = True
    
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "List Payment - Tanggal " & Format(tgl_mulai1.Value, "DD-MM-YYYY") & " Sampai " & Format(tgl_akhir1.Value, "DD-MM-YYYY")
        .Cells(1, 1).Font.Name = "Verdana"
        .Cells(1, 1).Font.Bold = True
        .Cells(4, 1).Value = "NO"
        .Cells(4, 2).Value = "CARD NUMBER"
        .Cells(4, 3).Value = "CH NAME"
        .Cells(4, 4).Value = "AGENT"
        .Cells(4, 5).Value = "PAYMENT"
        .Cells(4, 6).Value = "PAYDATE"
        .Cells(4, 7).Value = "PRODUCT"

        iRow = 4
        If rs.RecordCount > 0 Then
            ProgressBar1.Max = rs.RecordCount
            i = 0
            Do Until rs.EOF
                i = i + 1
                iRow = iRow + 1
                ProgressBar1.Value = rs.Bookmark
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 3).Value = IIf(IsNull(rs!Name), "", rs!Name)
                .Cells(iRow, 4).Value = IIf(IsNull(rs!agent), "", rs!agent)
                .Cells(iRow, 5).Value = Format(IIf(IsNull(rs!Payment), "", rs!Payment), "##,###")
                .Cells(iRow, 6).Value = IIf(IsNull(rs("paydate")), "", Format(rs("paydate"), "DD-MM-YYYY"))
                .Cells(iRow, 7).Value = IIf(IsNull(rs!acc_type), "", rs!acc_type)
                rs.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 14
            ExlObj.Cells(4, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
        ProgressBar1.Value = 0
        cmd_download.Enabled = True
    
        Set ExlObj = Nothing
        Set rs = Nothing

        'StartMeUp (Txtlocation.Text)
        'FILL COLOR CELL
        'ExlObj.Range(.Cells(NoUrut, 1), .Cells(NoUrut, 7)).Interior.Color = RGB(6, 207, 250)
    End With
End Sub

Private Sub Command2_Click()
On Error GoTo bawah
    Call IsiAccountPTP_ByTanggal
    Call HitungSeluruh
    Call IsiAccountPayment_ByTanggal
bawah:
End Sub

Private Sub IsiAccountPayment_ByTanggal()
    Dim listItem As listItem
    Dim agent As String
    Dim total_payment As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim jam_mulai As String
    Dim jam_selesai As String
    Dim tgl_mulai As String
    Dim tgl_akhir As String
    Dim tgl_akhir_server As Date
    Dim tgl_akhir_criteria As Date
    Dim tgl_mulai_criteria As Date
    On Error GoTo bawah
    
    Me.MousePointer = vbHourglass
    
    jam_mulai = "00:00:00"
    jam_selesai = "23:59:59"
    
    tanggal_sekarang = Format(tgl_mulai1.Value, "yyyy-mm-dd")
    
    bulan_sekarang = Format(tanggal_sekarang, "MM")
    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
    
    'LOCALHOST
'    tgl_mulai = Format(tgl_mulai1.Value, "MM-DD-YYYY")
'    tgl_akhir = Format(tgl_akhir1.Value, "MM-DD-YYYY")
    
    tgl_mulai = Format(tgl_mulai1.Value, "yyyy-mm-dd")
    tgl_akhir = Format(tgl_akhir1.Value, "yyyy-mm-dd")
    
    If Rs_list.state = 1 Then Rs_list.Close
    
    Rs_list.Open "SELECT MAX(paydate) as paydate FROM tbllunas"
    
    tgl_akhir_server = Format(IIf(IsNull(Rs_list!paydate), "1900-01-01", Rs_list!paydate), "yyyy-mm-dd")
    
    tgl_akhir_criteria = tgl_akhir
    tgl_mulai_criteria = tgl_mulai
    
    If tgl_akhir_criteria > tgl_akhir_server Then
        tgl_akhir = Format$(tgl_akhir_server, "yyyy-mm-dd")
    End If
    
    If tgl_mulai_criteria > tgl_akhir_server Then
        tgl_mulai = Format$(tgl_akhir_server, "yyyy-mm-dd")
    End If
    agent = cmb_agent.text
    
    If cmb_agent = " " Then
        agent = ""
    End If
    
    agent = cmb_agent.text
    
    If cmb_agent = " " Then
        agent = ""
    End If
    
    LvPayment.ListItems.clear
    If Rs_list.state = 1 Then Rs_list.Close
    
    If cmb_agent = "ALL" Then
    'in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))
'        Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
'                             "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
'                             "WHERE Payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' " & _
'                             "AND agent ilike 'D%' and char_length(agent) = 4) as a left join  " & _
'                             "(SELECT custid, name, acc_type from mgm where char_length(agent) = 4) as b on a.custid = b.custid "
        
'        If Left(MDIForm1.Text2.text, 2) = "AM" Then
'            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
'                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
'                         "WHERE Payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' " & _
'                         "AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as a left join  " & _
'                         "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as b on a.custid = b.custid "
'        Else
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
                             "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                             "WHERE Payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' " & _
                             "and char_length(agent) > 1) as a left join  " & _
                             "(SELECT custid, name, acc_type from mgm where char_length(agent) > 1) as b on a.custid = b.custid "
'        End If
    '                            "WHERE Payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' " & _

    Else
        If Check1.Value = 1 Then
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                         "WHERE Payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' " & _
                         "AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as a left join  " & _
                         "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as b on a.custid = b.custid "
        ElseIf f_team = False Then
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                         "WHERE Payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' " & _
                         "AND agent = '" & cmb_agent.text & "') as a left join  " & _
                         "(SELECT custid, name, acc_type from mgm where agent = '" & cmb_agent.text & "') as b on a.custid = b.custid "
        
        '"WHERE Payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' " & _

        ElseIf f_team = True Then
        'ubahtian left diganti inner
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM (" & _
                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                         "WHERE Payment > 100 AND date(paydate) between '" & tgl_mulai & "' AND '" & tgl_akhir & "' " & _
                         "AND agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%')) as a inner join  " & _
                         "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%')) as b on a.custid = b.custid "
         
         '"WHERE Payment > 100 AND paydate between '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
         End If
    End If
     
     
        If Rs_list.RecordCount > 0 Then
          Do Until Rs_list.EOF
              Set listItem = LvPayment.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", CStr(Rs_list!CustId)))
                              listItem.SubItems(1) = IIf(IsNull(Rs_list!agent), "", Rs_list!agent)
                              listItem.SubItems(2) = Format(IIf(IsNull(Rs_list!Payment), "", Rs_list!Payment), "##,###")
                              listItem.SubItems(3) = cnull(IIf(IsNull(Rs_list!paydate), "", Rs_list!paydate))
                              listItem.SubItems(4) = cnull(IIf(IsNull(Rs_list!ID), "", Rs_list!ID))
                              listItem.SubItems(5) = cnull(IIf(IsNull(Rs_list!Name), "", Rs_list!Name))
                              listItem.SubItems(6) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
                              
                              Total = Total + IIf(IsNull(Rs_list!Payment), "", Rs_list!Payment)
                              
              Rs_list.MoveNext
          Loop
          
          lblpayment.Caption = "TOTAL : IDR " & Format(Total, "##,###") & " "
          txt_total_payment.text = Total
        Else
          MsgBox "Data Payment Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
          lblpayment.Caption = "TOTAL : IDR 0 "
          txt_total_payment.text = 0
        End If
        
        Me.MousePointer = vbNormal
bawah:
End Sub


Private Sub IsiAccountPTP_ByTanggal()
    Dim listItem As listItem
    Dim agent As String
    Dim total_ptp As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim id_payment As String
    Dim jam_mulai As String
    Dim jam_selesai As String
    Dim tgl_mulai As String
    Dim tgl_akhir As String
    On Error GoTo bawah
    jam_mulai = "00:00:00"
    jam_selesai = "23:59:59"
    

    
    tanggal_sekarang = Format(tgl_mulai1.Value, "yyyy-mm-dd")
    
    bulan_sekarang = Format(tanggal_sekarang, "MM")
    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
    
    'LOCALHOST
'    tgl_mulai = Format(tgl_mulai1.Value, "MM-DD-YYYY")
'    tgl_akhir = Format(tgl_akhir1.Value, "MM-DD-YYYY")
    
    tgl_mulai = Format(tgl_mulai1.Value, "yyyy-mm-dd")
    tgl_akhir = Format(tgl_akhir1.Value, "yyyy-mm-dd")
    
    agent = cmb_agent.text
    
    If cmb_agent = " " Then
        agent = ""
    End If
    
    
    LvPTP.ListItems.clear
    If Rs_list.state = 1 Then Rs_list.Close
    
    If cmb_agent = "ALL" Then
'        If Left(MDIForm1.Text2.text, 2) = "AM" Then
'            Rs_list.Open "SELECT * FROM( " & _
'                         "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( " & _
'                         "SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
'                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
'                         "order by custid) As a RIGHT JOIN ( " & _
'                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
'                         "WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
'                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
'                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
'                         " order by custid_bayar DESC) a  " & _
'                         "where id in (" & _
'                         "select id from (" & _
'                         "SELECT b.custid, max(id) as id FROM( " & _
'                         "SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
'                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
'                         "order by custid) As a RIGHT JOIN ( " & _
'                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
'                         "WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
'                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
'                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
'                         " group by 1 order by 1 DESC ) a " & _
'                         " ) "
'        Else
            Rs_list.Open "select * FROM( " & _
                    "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name, acc_type FROM( " & _
                    "SELECT Distinct custid FROM tbllunas WHERE agent ilike 'D%' and char_length(agent) = 4 " & _
                    "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                    "order by custid) As a RIGHT JOIN ( " & _
                    "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                    "WHERE mgm.custid = tblnegoptp.custid AND agent ilike 'D%' and char_length(agent) = 4 " & _
                    "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                    "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid order by custid_bayar DESC ) a " & _
                    "where id in (" & _
                    "select id from ( " & _
                    "SELECT b.custid, max(id) as id FROM( " & _
                    "SELECT Distinct custid FROM tbllunas WHERE agent ilike 'D%' and char_length(agent) = 4 " & _
                    "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                    "order by custid) As a RIGHT JOIN ( " & _
                    "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                    "WHERE mgm.custid = tblnegoptp.custid AND agent ilike 'D%' and char_length(agent) = 4 " & _
                    "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                    "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
                    " group by 1 order by 1 DESC) a " & _
                    " ) "
'        End If

'                    "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name, acc_type FROM( " & _
'                    "SELECT Distinct custid FROM tbllunas WHERE agent ilike 'D%' and char_length(agent) = 4 " & _
'                    "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
'                    "order by custid) As a RIGHT JOIN ( " & _
'                    "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
'                    "WHERE mgm.custid = tblnegoptp.custid AND agent ilike 'D%' and char_length(agent) = 4 " & _
'                    "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
'                    "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
'                    " order by custid_bayar DESC "
                    
                    '"AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' " & _
                    '"AND '" & tgl_akhir & " " & jam_selesai & "' ) As b on a.custid = b.custid order by custid_bayar DESC "

    Else
        If Check1.Value = 1 Then
                    Rs_list.Open "SELECT * FROM( " & _
                         "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                         "order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
                         " order by custid_bayar DESC) a  " & _
                         "where id in (" & _
                         "select id from (" & _
                         "SELECT b.custid, max(id) as id FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                         "order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) " & _
                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
                         " group by 1 order by 1 DESC ) a " & _
                         " ) "
        ElseIf f_team = False Then
            Rs_list.Open "SELECT * FROM( " & _
                         "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent = '" & agent & "' " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                         "order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent = '" & agent & "' " & _
                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
                         " order by custid_bayar DESC) a  " & _
                         "where id in (" & _
                         "select id from (" & _
                         "SELECT b.custid, max(id) as id FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent = '" & agent & "' " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                         "order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent = '" & agent & "' " & _
                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid " & _
                         " group by 1 order by 1 DESC ) a " & _
                         " ) "
        
                    '    "AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' " & _
                    '    "AND '" & tgl_akhir & " " & jam_selesai & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
        
        
        ElseIf f_team = True Then
            Rs_list.Open "SELECT * FROM( " & _
                         "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                         " order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') " & _
                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid order by custid_bayar DESC ) a " & _
                         " where id in ( " & _
                         " select id from  (" & _
                         " select b.custid, max(id) as id  FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                         " order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') " & _
                         "AND promisepay > 100 AND date(promisedate) between '" & tgl_mulai & "' " & _
                         "AND '" & tgl_akhir & "' ) As b on a.custid = b.custid group by 1 order by 1 DESC ) a" & _
                         " ) "
                    
                    '     "AND promisepay > 100 AND promisedate between '" & tgl_mulai & " " & jam_mulai & "' " & _
                    '     "AND '" & tgl_akhir & " " & jam_selesai & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
        End If
    End If
        
'    Rs_list.Open "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay FROM mgm, tblnegoptp  " & _
'                 "WHERE mgm.custid = tblnegoptp.custid " & _
'                 "AND agent = '" & agent & "' AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
'                 "AND date_part('year',promisedate) = '" & tahun_sekarang & "' order by promisedate desc "
    
    If Rs_list.RecordCount > 0 Then
          Do Until Rs_list.EOF
              Set listItem = LvPTP.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", CStr(Rs_list!CustId)))
                              listItem.SubItems(1) = IIf(IsNull(Rs_list!agent), "", Rs_list!agent)
                              listItem.SubItems(2) = Format(IIf(IsNull(Rs_list!PromisePay), "", Rs_list!PromisePay), "##,###")
                              listItem.SubItems(3) = cnull(IIf(IsNull(Rs_list!PromiseDate), "", Rs_list!PromiseDate))
                              listItem.SubItems(4) = cnull(IIf(IsNull(Rs_list!ID), "", Rs_list!ID))
                              listItem.SubItems(5) = IIf(IsNull(Rs_list!custid_bayar), "", Rs_list!custid_bayar)
                              listItem.SubItems(6) = IIf(IsNull(Rs_list!Name), "", Rs_list!Name)
                              listItem.SubItems(7) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
                              
                              Total = Total + IIf(IsNull(Rs_list!PromisePay), "", Rs_list!PromisePay)
                              
                              id_payment = IIf(IsNull(Rs_list!custid_bayar), "", Rs_list!custid_bayar)

                              If id_payment <> "" Then
                                    For K = 1 To 7
                                          'LvPTP.ListItems(Rs_list.Bookmark).ListSubItems(K).ForeColor = vbBlue
                                          listItem.ListSubItems(K).ForeColor = vbBlue
                                          listItem.ForeColor = vbBlue
                                    Next K
                              End If
              Rs_list.MoveNext
          Loop
          lbldata.Caption = "Jumlah Data  : " & Rs_list.RecordCount & " Rows"
          lbltotal.Caption = "TOTAL : IDR " & Format(Total, "##,###") & " "
          txt_total_ptp.text = Total
      Else
          MsgBox "Data Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
          LvPayment.ListItems.clear
          lbldata.Caption = "Rows : 0"
          lbltotal.Caption = "TOTAL : IDR 0 "
          txt_total_ptp.text = 0
      End If
bawah:
End Sub






Private Sub Command3_Click()

End Sub

Private Sub LvPTP_ColumnClick(ByVal ColumnHeader As ColumnHeader)
On Error GoTo bawah
   LvPTP.SortKey = ColumnHeader.Index - 1
   IndexColumnHEader = ColumnHeader.Index - 1
   LvPTP.Sorted = True
bawah:
End Sub

Private Sub cmd_paid_Click()
    Dim w As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim cmdsql As String
    Dim tanggal_bayar As String
    Dim id_ptp1 As String
    Dim CustId As String
    On Error GoTo bawah

    
    If LvPayment.ListItems.Count = 0 Then
        MsgBox "Tidak Ada Payment Untuk PTP Tersebut !", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    a = MsgBox("Apakah Anda Yakin Akan Merubah Tanggal PTP?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses Dibatalkan!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek > 1 Then
        MsgBox "Anda Tidak Boleh Memilih Lebih Dari 1 PTP!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    If cek = 0 Then
        MsgBox "Check PTP Terlebih Dahulu!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    tanggal_bayar = Format$(LvPayment.SelectedItem.SubItems(3), "YYYY-MM-DD")
    id_ptp1 = LvPTP.SelectedItem.SubItems(4)
    CustId = LvPTP.SelectedItem.text
    
    For w = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(w).Checked = True Then
            cmdsql = "UPDATE tblnegoptp SET promisedate = '" & tanggal_bayar & "' "
            cmdsql = cmdsql + " WHERE id = '" & id_ptp1 & "' "
            M_OBJCONN.Execute cmdsql
            
            cmdsql = "UPDATE mgm SET dateptp = '" & tanggal_bayar & "' "
            cmdsql = cmdsql + " WHERE custid = '" & CustId & "' "
            M_OBJCONN.Execute cmdsql
        End If
    Next w
    
    MsgBox "Tanggal PTP Berhasil Di-Update!", vbOKOnly + vbInformation, "Informasi"
    LvPTP.ListItems(1).Checked = False
    Call IsiAccountPTP
bawah:
End Sub

Private Sub cmd_payment_Click()
    Call IsiAccountBayar
    Call HitungSeluruh
End Sub

Private Sub cmd_showptp_Click()
On Error GoTo bawah
    Call IsiAccountPTP
    Call HitungSeluruh
bawah:
End Sub

Private Sub HitungSeluruh()
    Dim TotalPtp As Double
    Dim TotalPayment As Double
    Dim totalseluruh As Double
    On Error GoTo HELL
    
    TotalPtp = txt_total_ptp.text
    TotalPayment = txt_total_payment.text
    
    totalseluruh = TotalPtp + TotalPayment

    lbl_total_keseluruhan.Caption = "CONFIDENT : IDR " + Format(totalseluruh, "##,###")
HELL:
    'MsgBox err.Description
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Call koneksi
    Call IsiAgent
    Call HeaderListPTP
    Call HeaderListPayment
    
    If MDIForm1.Text2.text = "Supervisor" Or MDIForm1.Text2.text = "Manager" Or Left(MDIForm1.Text2.text, 2) = "AM" Then
        Frame3.Visible = True
    End If
    
    If Left(MDIForm1.Text2.text, 2) <> "AM" Then
        Check1.Visible = False
    End If
    
    f_team = False
    
    Me.Width = Screen.Width - 500
    'msfx.Width = Screen.Width - 500
    
    DTPicker1.Value = Now
    tgl_mulai1.Value = Now
    tgl_akhir1.Value = Now
    
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        lbl_team.Visible = False
        cmb_team.Visible = False
        chk_team.Visible = False
        txt_date.Visible = False
        cmd_change.Visible = False
        txt_amount.Visible = False
        cmd_amount.Visible = False
        cmd_download.Visible = False
        cmd_download_payment.Visible = False
    Else
        If Rs_list.state = 1 Then Rs_list.Close

        If Left(MDIForm1.Text1.text, 2) = "TL" Then
            Rs_list.Open "select userid from usertbl where usertype = '1' AND userid ilike 'D%' AND  team = '" & MDIForm1.Text1.text & "' Order by userid"
        Else
            Rs_list.Open "SELECT DISTINCT team from usertbl WHERE team ilike 'TL%'"
        End If

    End If
End Sub

Private Sub IsiAgent()
    If Rs_list.state = 1 Then Rs_list.Close
    
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        Rs_list.Open "select userid from usertbl where usertype = '1' and userid = '" & MDIForm1.Text1.text & "' Order by userid"
    ElseIf UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        Rs_list.Open "select userid from usertbl where usertype = '1' AND userid ilike 'D%' AND  team = '" & MDIForm1.Text1.text & "' Order by userid"
    'ElseIf Left(MDIForm1.Text2.text, 2) = "AM" Then
    '    Rs_list.Open "select userid from usertbl where usertype = '1' and userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "') Order by userid"
    Else
        Rs_list.Open "select userid from usertbl where usertype = '1' and userid ilike 'D%' Order by userid"
    End If
    
    cmb_agent.AddItem " "
    
    If UCase(MDIForm1.Text2.text) <> "AGENT" And UCase(MDIForm1.Text2.text) <> "TEAMLEADER" Then
        cmb_agent.AddItem "ALL"
    End If
    
    
    While Not Rs_list.EOF
        cmb_agent.AddItem Rs_list("USERID")
        Rs_list.MoveNext
    Wend
End Sub

Private Sub HeaderListPTP()
    LvPTP.ColumnHeaders.ADD , , "Custid", 2100
    LvPTP.ColumnHeaders.ADD , , "Agent", 1000
    LvPTP.ColumnHeaders.ADD , , "PromisePay", 1300
    LvPTP.ColumnHeaders.ADD , , "PromiseDate", 1300
    LvPTP.ColumnHeaders.ADD , , "ID", 0
    LvPTP.ColumnHeaders.ADD , , "Custid Bayar", 0
    LvPTP.ColumnHeaders.ADD , , "CH Name", 2350
    LvPTP.ColumnHeaders.ADD , , "Product", 1500
End Sub

Private Sub HeaderListPayment()
    LvPayment.ColumnHeaders.ADD , , "Custid", 2100
    LvPayment.ColumnHeaders.ADD , , "Agent", 1000
    LvPayment.ColumnHeaders.ADD , , "Payment", 1300
    LvPayment.ColumnHeaders.ADD , , "Paydate", 1300
    LvPayment.ColumnHeaders.ADD , , "ID", 0
    LvPayment.ColumnHeaders.ADD , , "CH Name", 2350
    LvPayment.ColumnHeaders.ADD , , "Product", 1500
End Sub

Private Sub IsiAccountPTP()
    Dim listItem As listItem
    Dim agent As String
    Dim total_ptp As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim id_payment As String
    Dim TotalKeseluruhan As Double
    On Error GoTo HELL
    
    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
    
    bulan_sekarang = Format(tanggal_sekarang, "MM")
    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
    
    agent = cmb_agent.text
    
    If cmb_agent = " " Then
        agent = ""
    End If
    
    LvPTP.ListItems.clear
    If Rs_list.state = 1 Then Rs_list.Close
    
    If cmb_agent = "ALL" Then
'        If Left(MDIForm1.Text2.text, 2) = "AM" Then
'            Rs_list.Open "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name, acc_type FROM( " & _
'                    "SELECT Distinct custid FROM tbllunas WHERE  agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) and char_length(agent) > 1 " & _
'                    "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
'                    "AND date_part('year',paydate) = '" & tahun_sekarang & "' order by custid) As a RIGHT JOIN ( " & _
'                    "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
'                    "WHERE mgm.custid = tblnegoptp.custid  and char_length(agent) > 1 " & _
'                    "AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
'                    "AND date_part('year',promisedate) = '" & tahun_sekarang & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
'        Else
            Rs_list.Open "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name, acc_type FROM( " & _
                        "SELECT Distinct custid FROM tbllunas WHERE char_length(agent) > 1 " & _
                        "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                        "AND date_part('year',paydate) = '" & tahun_sekarang & "' order by custid) As a RIGHT JOIN ( " & _
                        "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                        "WHERE mgm.custid = tblnegoptp.custid  and char_length(agent) > 1 " & _
                        "AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
                        "AND date_part('year',promisedate) = '" & tahun_sekarang & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
'        End If
    Else
        If Check1.Value = 1 Then
            Rs_list.Open "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name, acc_type FROM( " & _
                    "SELECT Distinct custid FROM tbllunas WHERE  agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "')) and char_length(agent) > 1 " & _
                    "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                    "AND date_part('year',paydate) = '" & tahun_sekarang & "' order by custid) As a RIGHT JOIN ( " & _
                    "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                    "WHERE mgm.custid = tblnegoptp.custid  and char_length(agent) > 1 " & _
                    "AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
                    "AND date_part('year',promisedate) = '" & tahun_sekarang & "' ) As b on a.custid = b.custid order by custid_bayar DESC "

        ElseIf f_team = False Then
            Rs_list.Open "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent = '" & agent & "' " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                         "AND date_part('year',paydate) = '" & tahun_sekarang & "' order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent = '" & agent & "' " & _
                         "AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
                         "AND date_part('year',promisedate) = '" & tahun_sekarang & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
        ElseIf f_team = True Then
            Rs_list.Open "SELECT coalesce(a.custid,'') as custid_bayar, id, agent, b.custid, promisedate, promisepay, name,acc_type FROM( " & _
                         "SELECT Distinct custid FROM tbllunas WHERE agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') " & _
                         "AND payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                         "AND date_part('year',paydate) = '" & tahun_sekarang & "' order by custid) As a RIGHT JOIN ( " & _
                         "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay, name, acc_type FROM mgm, tblnegoptp " & _
                         "WHERE mgm.custid = tblnegoptp.custid AND agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%') " & _
                         "AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
                         "AND date_part('year',promisedate) = '" & tahun_sekarang & "' ) As b on a.custid = b.custid order by custid_bayar DESC "
        End If
    End If
        
'    Rs_list.Open "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay FROM mgm, tblnegoptp  " & _
'                 "WHERE mgm.custid = tblnegoptp.custid " & _
'                 "AND agent = '" & agent & "' AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
'                 "AND date_part('year',promisedate) = '" & tahun_sekarang & "' order by promisedate desc "
    
    If Rs_list.RecordCount > 0 Then
          Do Until Rs_list.EOF
              Set listItem = LvPTP.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", CStr(Rs_list!CustId)))
                              listItem.SubItems(1) = IIf(IsNull(Rs_list!agent), "", Rs_list!agent)
                              listItem.SubItems(2) = Format(IIf(IsNull(Rs_list!PromisePay), "", Rs_list!PromisePay), "##,###")
                              listItem.SubItems(3) = cnull(IIf(IsNull(Rs_list!PromiseDate), "", Rs_list!PromiseDate))
                              listItem.SubItems(4) = cnull(IIf(IsNull(Rs_list!ID), "", Rs_list!ID))
                              listItem.SubItems(5) = IIf(IsNull(Rs_list!custid_bayar), "", Rs_list!custid_bayar)
                              listItem.SubItems(6) = IIf(IsNull(Rs_list!Name), "", Rs_list!Name)
                              listItem.SubItems(7) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
                              
                              Total = Total + IIf(IsNull(Rs_list!PromisePay), "", Rs_list!PromisePay)
                              
                              id_payment = IIf(IsNull(Rs_list!custid_bayar), "", Rs_list!custid_bayar)

                              If id_payment <> "" Then
                                    For K = 1 To 7
                                          'LvPTP.ListItems(Rs_list.Bookmark).ListSubItems(K).ForeColor = vbBlue
                                          listItem.ListSubItems(K).ForeColor = vbBlue
                                          listItem.ForeColor = vbBlue
                                    Next K
                              End If
              Rs_list.MoveNext
          Loop
          lbldata.Caption = "Jumlah Data  : " & Rs_list.RecordCount & " Rows"
          lbltotal.Caption = "TOTAL : IDR " & Format(Total, "##,###") & " "
          txt_total_ptp.text = Total
      Else
          MsgBox "Data Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
          
          LvPayment.ListItems.clear
          lbldata.Caption = "Rows : 0"
          lbltotal.Caption = "TOTAL : IDR 0 "
          txt_total_ptp.text = 0
      End If
HELL:
'      MsgBox err.Description
End Sub


Private Sub IsiAccountBayar()
    Dim listItem As listItem
    Dim agent As String
    Dim total_payment As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    On Error GoTo bawah
    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
    
    bulan_sekarang = Format(tanggal_sekarang, "MM")
    tahun_sekarang = Format(tanggal_sekarang, "YYYY")

    agent = cmb_agent.text
    
    If cmb_agent = " " Then
        agent = ""
    End If
    
    LvPayment.ListItems.clear
    If Rs_list.state = 1 Then Rs_list.Close
    
    If cmb_agent = "ALL" Then
'        If Left(MDIForm1.Text2.text, 2) = "AM" Then
'            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
'                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
'                         "WHERE Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
'                         "AND date_part('year',paydate) = '" & tahun_sekarang & "'  " & _
'                         "AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as a left join  " & _
'                         "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as b on a.custid = b.custid "
'        Else
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
                                 "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                                 "WHERE Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                                 "AND date_part('year',paydate) = '" & tahun_sekarang & "') as a left join  " & _
                                 "(SELECT custid, name, acc_type from mgm where agent ilike 'D%' and char_length(agent) = 4) as b on a.custid = b.custid "
'        End If
    Else
        If Check1.Value = 1 Then
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                         "WHERE Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                         "AND date_part('year',paydate) = '" & tahun_sekarang & "'  " & _
                         "AND agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as a left join  " & _
                         "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'))) as b on a.custid = b.custid "

        ElseIf f_team = False Then
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM ( " & _
                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                         "WHERE Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                         "AND date_part('year',paydate) = '" & tahun_sekarang & "'  " & _
                         "AND agent = '" & cmb_agent.text & "') as a left join  " & _
                         "(SELECT custid, name, acc_type from mgm where agent = '" & cmb_agent.text & "') as b on a.custid = b.custid "
        ElseIf f_team = True Then
            Rs_list.Open "SELECT a.custid as custid, paydate, payment, agent, name, a.id, acc_type FROM (" & _
                         "SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                         "WHERE Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                         "AND date_part('year',paydate) = '" & tahun_sekarang & "'  " & _
                         "AND agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%')) as a left join  " & _
                         "(SELECT custid, name, acc_type from mgm where agent in (select userid from usertbl where team = '" & cmb_team.text & "' AND userid ilike  'D%')) as b on a.custid = b.custid "
        End If
    End If
     
     
        If Rs_list.RecordCount > 0 Then
          Do Until Rs_list.EOF
              Set listItem = LvPayment.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", CStr(Rs_list!CustId)))
                              listItem.SubItems(1) = IIf(IsNull(Rs_list!agent), "", Rs_list!agent)
                              listItem.SubItems(2) = Format(IIf(IsNull(Rs_list!Payment), "", Rs_list!Payment), "##,###")
                              listItem.SubItems(3) = cnull(IIf(IsNull(Rs_list!paydate), "", Rs_list!paydate))
                              listItem.SubItems(4) = cnull(IIf(IsNull(Rs_list!ID), "", Rs_list!ID))
                              listItem.SubItems(5) = cnull(IIf(IsNull(Rs_list!Name), "", Rs_list!Name))
                              listItem.SubItems(6) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
                              
                              Total = Total + IIf(IsNull(Rs_list!Payment), "", Rs_list!Payment)
                              
              Rs_list.MoveNext
          Loop
          
          lblpayment.Caption = "TOTAL : IDR " & Format(Total, "##,###") & " "
          txt_total_payment.text = Total
        Else
          MsgBox "Data Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
          lblpayment.Caption = "TOTAL : IDR 0 "
          txt_total_payment.text = 0
        End If
bawah:
End Sub

Private Sub LvPayment_DblClick()
    If LvPayment.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = LvPayment.SelectedItem.text
        Form_ptp_payment.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub LvPayment_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Call Hitung_Payment_Dipilih
End Sub

Private Sub Hitung_Payment_Dipilih()
    
    Dim w As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim janji_bayar As Double
    Dim total_janji_bayar As Double
    On Error GoTo bawah
    
    cek = 0

    For K = 1 To LvPayment.ListItems.Count
        If LvPayment.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K

    For w = 1 To LvPayment.ListItems.Count
        If LvPayment.ListItems(w).Checked = True Then
            janji_bayar = LvPayment.ListItems(w).SubItems(2)

            total_janji_bayar = total_janji_bayar + janji_bayar
        End If
    Next w

    lbl_total_hitung_payment.Caption = "TOTAL : IDR " + Format(total_janji_bayar, "##,###")
    lbl_jumlah_data_dipilih_payment.Caption = "Jumlah " & cek & " Rows"
bawah:
End Sub


Private Sub LvPayment_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   LvPayment.SortKey = ColumnHeader.Index - 1
   IndexColumnHEader = ColumnHeader.Index - 1
   LvPayment.Sorted = True
End Sub



Private Sub Isi_payment()
    Dim listItem As listItem
    Dim agent As String
    Dim total_ptp As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    On Error GoTo bawah
        
    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
    
    bulan_sekarang = Format(tanggal_sekarang, "MM")
    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
'    bulan_sekarang = "06"
'    tahun_sekarang = "2015"
    
    LvPayment.ListItems.clear
    If Rs_list.state = 1 Then Rs_list.Close

    Rs_list.Open "SELECT * FROM ( " & _
                 "(SELECT custid, paydate, payment, agent, id FROM tbllunas  " & _
                 "WHERE custid = '" & LvPTP.SelectedItem.text & "' " & _
                 "AND Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                 "AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                 "AND paydate = (SELECT MAX(paydate) FROM tbllunas WHERE custid = '" & LvPTP.SelectedItem.text & "' " & _
                 "AND Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                 "AND date_part('year',paydate) = '" & tahun_sekarang & "')) AS a " & _
                 "LEFT JOIN (SELECT custid, acc_type, name FROM mgm where custid =  '" & LvPTP.SelectedItem.text & "') AS b on a.custid = b.custid)"

    If Rs_list.RecordCount > 0 Then
          Do Until Rs_list.EOF
              Set listItem = LvPayment.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", CStr(Rs_list!CustId)))
                              listItem.SubItems(1) = IIf(IsNull(Rs_list!agent), "", Rs_list!agent)
                              listItem.SubItems(2) = Format(IIf(IsNull(Rs_list!Payment), "", Rs_list!Payment), "##,###")
                              listItem.SubItems(3) = cnull(IIf(IsNull(Rs_list!paydate), "", Rs_list!paydate))
                              listItem.SubItems(4) = cnull(IIf(IsNull(Rs_list!ID), "", Rs_list!ID))
                              listItem.SubItems(5) = cnull(IIf(IsNull(Rs_list!Name), "", Rs_list!Name))
                              listItem.SubItems(6) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
              Rs_list.MoveNext
          Loop
    End If
bawah:
End Sub

Private Sub LvPTP_Click()
    If LvPTP.ListItems.Count <= 0 Then
        MsgBox "Tampilkan Data Terlebih Dahulu !", vbOKOnly + vbInformation, "Perhatian"
    Exit Sub
    End If

    Call Isi_payment
End Sub

Private Sub LvPTP_DblClick()
    If LvPTP.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = LvPTP.SelectedItem.text
        Form_ptp_payment.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub LvPTP_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Call cmd_count_Click
End Sub

Private Sub txt_amount_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub


