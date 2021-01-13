VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Form_Distribute_Otomatis 
   BackColor       =   &H8000000E&
   Caption         =   "Distibute Otomatis"
   ClientHeight    =   10470
   ClientLeft      =   3780
   ClientTop       =   645
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   14055
   Begin VB.Frame Frame4 
      Height          =   9765
      Left            =   0
      TabIndex        =   3
      Top             =   705
      Width           =   13935
      Begin VB.ComboBox cmbflag 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form_Distribute_Otomatis.frx":0000
         Left            =   4305
         List            =   "Form_Distribute_Otomatis.frx":0002
         TabIndex        =   62
         Top             =   4620
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.ComboBox Cmb_ShowUser 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   555
         TabIndex        =   60
         Top             =   4635
         Width           =   2985
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   390
         Left            =   6945
         TabIndex        =   59
         Top             =   6390
         Width           =   375
      End
      Begin VB.CommandButton cmdLoadUser 
         Caption         =   "Refresh User"
         Height          =   390
         Left            =   195
         TabIndex        =   56
         Top             =   9300
         Width           =   1155
      End
      Begin VB.CommandButton cmdException 
         Caption         =   ">>"
         Height          =   390
         Left            =   6945
         TabIndex        =   36
         Top             =   5880
         Width           =   375
      End
      Begin VB.Frame frmUserAssign 
         Height          =   1305
         Left            =   1125
         TabIndex        =   35
         Top             =   5760
         Visible         =   0   'False
         Width           =   4815
         Begin Threed.SSCommand cmdCancel2 
            Height          =   345
            Left            =   4470
            TabIndex        =   55
            Top             =   -15
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   609
            _Version        =   196610
            ForeColor       =   64
            BackColor       =   8421631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "X"
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Height          =   615
            Left            =   3855
            TabIndex        =   40
            Top             =   360
            Width           =   705
         End
         Begin VB.TextBox txtUserAssign 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1455
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   375
            Width           =   2325
         End
         Begin TDBNumber6Ctl.TDBNumber txtJmlAssign 
            Height          =   330
            Left            =   1455
            TabIndex        =   53
            Top             =   720
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   582
            Calculator      =   "Form_Distribute_Otomatis.frx":0004
            Caption         =   "Form_Distribute_Otomatis.frx":0024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_Distribute_Otomatis.frx":0090
            Keys            =   "Form_Distribute_Otomatis.frx":00AE
            Spin            =   "Form_Distribute_Otomatis.frx":00F8
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
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
            MaxValue        =   9999999999999
            MinValue        =   -99999999999
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
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin VB.Label Label11 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "User :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   675
            TabIndex        =   39
            Top             =   390
            Width           =   630
         End
         Begin VB.Label Label10 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   615
            TabIndex        =   37
            Top             =   750
            Width           =   630
         End
      End
      Begin VB.CommandButton cmdProsesDistribusi 
         Caption         =   "Proses Distribusi"
         Height          =   390
         Left            =   5445
         TabIndex        =   34
         Top             =   9300
         Width           =   1440
      End
      Begin VB.Frame Frame6 
         Height          =   825
         Left            =   6990
         TabIndex        =   28
         Top             =   4170
         Width           =   6870
         Begin VB.TextBox txtTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3900
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "0"
            Top             =   465
            Width           =   780
         End
         Begin VB.TextBox txtTotalDtaCek 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   5565
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "0"
            Top             =   450
            Width           =   780
         End
         Begin TDBNumber6Ctl.TDBNumber tdbDataToDistribute 
            Height          =   330
            Left            =   2040
            TabIndex        =   58
            Top             =   435
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   582
            Calculator      =   "Form_Distribute_Otomatis.frx":0120
            Caption         =   "Form_Distribute_Otomatis.frx":0140
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_Distribute_Otomatis.frx":01AC
            Keys            =   "Form_Distribute_Otomatis.frx":01CA
            Spin            =   "Form_Distribute_Otomatis.frx":0214
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
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
            MaxValue        =   9999999999999
            MinValue        =   -99999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1761280005
            Value           =   0
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin TDBNumber6Ctl.TDBNumber tdbDatarata 
            Height          =   330
            Left            =   345
            TabIndex        =   66
            Top             =   435
            Visible         =   0   'False
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   582
            Calculator      =   "Form_Distribute_Otomatis.frx":023C
            Caption         =   "Form_Distribute_Otomatis.frx":025C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_Distribute_Otomatis.frx":02C8
            Keys            =   "Form_Distribute_Otomatis.frx":02E6
            Spin            =   "Form_Distribute_Otomatis.frx":0330
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            MaxValue        =   9999999999999
            MinValue        =   -99999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin VB.Label Label18 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Average Data/SPV :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   67
            Top             =   135
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label14 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Data To Distribute :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   57
            Top             =   135
            Width           =   1575
         End
         Begin VB.Label Label15 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total data Selected:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5160
            TabIndex        =   47
            Top             =   135
            Width           =   1755
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total data Remain:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3450
            TabIndex        =   29
            Top             =   135
            Width           =   1545
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   1740
         TabIndex        =   25
         Top             =   4185
         Width           =   2130
         Begin VB.OptionButton optUneven 
            Appearance      =   0  'Flat
            Caption         =   "UnEven"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   27
            Top             =   180
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optEven 
            Appearance      =   0  'Flat
            Caption         =   "Even"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   75
            TabIndex        =   26
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Manager"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4200
         Left            =   120
         TabIndex        =   16
         Top             =   5055
         Width           =   6855
         Begin VB.TextBox txttotalspv 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "0"
            Top             =   3840
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.CheckBox chkShowAssign 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Assign"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1125
            TabIndex        =   52
            Top             =   3765
            Width           =   1320
         End
         Begin VB.TextBox txtTotalGetData 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3570
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "0"
            Top             =   3840
            Width           =   780
         End
         Begin VB.TextBox txtTotaltUser 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   6255
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "0"
            Top             =   3840
            Width           =   525
         End
         Begin VB.CheckBox CheckAll_Manager 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            TabIndex        =   17
            Top             =   3765
            Visible         =   0   'False
            Width           =   1065
         End
         Begin MSComctlLib.ListView List_User 
            Height          =   3405
            Left            =   30
            TabIndex        =   18
            Top             =   315
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   6006
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
            Appearance      =   0
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
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   2760
            TabIndex        =   19
            Top             =   4440
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label17 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total SPV :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4590
            TabIndex        =   65
            Top             =   3705
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label16 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Distribusi :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2610
            TabIndex        =   51
            Top             =   3705
            Width           =   1020
         End
         Begin VB.Label Label8 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total User :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5745
            TabIndex        =   31
            Top             =   3705
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Distribute"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   90
            TabIndex        =   20
            Top             =   45
            Width           =   1245
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   0
            Top             =   45
            Width           =   6795
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Manager"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4275
         Left            =   75
         TabIndex        =   4
         Top             =   240
         Width           =   13620
         Begin VB.TextBox txtJmlCampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   12795
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "0"
            Top             =   1125
            Width           =   780
         End
         Begin VB.CommandButton cmdLoadDataCampaign 
            Caption         =   "Load Data"
            Height          =   390
            Left            =   4290
            TabIndex        =   24
            Top             =   1035
            Width           =   1440
         End
         Begin VB.ComboBox cmbCriteria 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   705
            Width           =   4395
         End
         Begin MSComctlLib.ListView List_Campaign 
            Height          =   2490
            Left            =   45
            TabIndex        =   8
            Top             =   1440
            Width           =   13515
            _ExtentX        =   23839
            _ExtentY        =   4392
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
            Appearance      =   0
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
         Begin VB.CheckBox CheckAll_Campaign 
            Appearance      =   0  'Flat
            Caption         =   "Check All"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   105
            TabIndex        =   7
            Top             =   3975
            Width           =   1455
         End
         Begin VB.ComboBox Cmb_BaseProduct 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "Form_Distribute_Otomatis.frx":0358
            Left            =   1350
            List            =   "Form_Distribute_Otomatis.frx":035A
            TabIndex        =   6
            Top             =   330
            Width           =   4395
         End
         Begin VB.TextBox Txt_Search 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1350
            TabIndex        =   5
            Top             =   1050
            Width           =   2850
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   2760
            TabIndex        =   9
            Top             =   4440
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label9 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Jml Campaign:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11415
            TabIndex        =   33
            Top             =   1140
            Width           =   1425
         End
         Begin VB.Label Label6 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Contain : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   23
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label Label5 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Kriteria : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   22
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Campaign"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   15
            Width           =   825
         End
         Begin VB.Label Label23 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Base Product"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   105
            Top             =   0
            Width           =   13470
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Manager"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2085
         Left            =   7095
         TabIndex        =   12
         Top             =   5070
         Width           =   6735
         Begin VB.Frame frmConfirmExeption 
            Height          =   1305
            Left            =   555
            TabIndex        =   41
            Top             =   420
            Visible         =   0   'False
            Width           =   5490
            Begin Threed.SSCommand cmdCancel 
               Height          =   345
               Left            =   5190
               TabIndex        =   54
               Top             =   0
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   609
               _Version        =   196610
               ForeColor       =   64
               BackColor       =   8421631
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "X"
            End
            Begin VB.TextBox txtReasonException 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1440
               TabIndex        =   44
               Top             =   570
               Width           =   2325
            End
            Begin VB.TextBox txtUserException 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   300
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   225
               Width           =   2325
            End
            Begin VB.CommandButton cmdOKException 
               Caption         =   "OK"
               Height          =   615
               Left            =   3885
               TabIndex        =   42
               Top             =   240
               Width           =   705
            End
            Begin VB.Label Label13 
               BackColor       =   &H00F1E5DB&
               BackStyle       =   0  'Transparent
               Caption         =   "Reason :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   615
               TabIndex        =   46
               Top             =   600
               Width           =   630
            End
            Begin VB.Label Label12 
               BackColor       =   &H00F1E5DB&
               BackStyle       =   0  'Transparent
               Caption         =   "User :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   675
               TabIndex        =   45
               Top             =   240
               Width           =   630
            End
         End
         Begin VB.CheckBox CheckAll_ProsesData 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            TabIndex        =   13
            Top             =   3375
            Width           =   1455
         End
         Begin MSComctlLib.ListView List_Exception 
            Height          =   1605
            Left            =   255
            TabIndex        =   14
            Top             =   315
            Width           =   6420
            _ExtentX        =   11324
            _ExtentY        =   2831
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
            Appearance      =   0
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exception"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   315
            TabIndex        =   15
            Top             =   15
            Width           =   960
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   240
            Top             =   15
            Width           =   6615
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1980
         Left            =   6990
         TabIndex        =   68
         Top             =   7515
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3493
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
         Appearance      =   0
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pembagian SPV"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7020
         TabIndex        =   69
         Top             =   7245
         Width           =   1260
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   6990
         Top             =   7245
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Flag"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3795
         TabIndex        =   63
         Top             =   4680
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   26
         Left            =   165
         TabIndex        =   61
         Top             =   4665
         Width           =   375
      End
   End
   Begin VB.CommandButton Cmd_Load 
      BackColor       =   &H00F1E5DB&
      Caption         =   "&Load Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5580
      Width           =   1830
   End
   Begin VB.CommandButton Cmd_Proses 
      BackColor       =   &H00F1E5DB&
      Caption         =   "&Proses Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5580
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   1
      Left            =   240
      Picture         =   "Form_Distribute_Otomatis.frx":035C
      Stretch         =   -1  'True
      Top             =   90
      Width           =   420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Distribute Otomatis"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   4665
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   -2250
      Picture         =   "Form_Distribute_Otomatis.frx":0E66
      Stretch         =   -1  'True
      Top             =   -150
      Width           =   20700
   End
End
Attribute VB_Name = "Form_Distribute_Otomatis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim sUser As String
Dim sLevel As String
Dim sExcepTrans As Boolean
Dim isiAwalShowUser As String
Private Sub CheckAll_Campaign_Click()
    Call CheckAll(List_Campaign, CheckAll_Campaign)
    Call itungJmlCek
    Call bagiRata
End Sub

Private Sub CheckAll_Manager_Click()
    Call CheckAll(List_User, CheckAll_Manager)
    Call bagiRata
End Sub

'Private Sub CheckAll_ProsesData_Click()
'    Call CheckAll(List_Proses_Data, CheckAll_ProsesData)
'End Sub

Private Sub Cmb_BaseProduct_Change()
    'Cmb_BaseProduct_Click
End Sub

Private Sub Cmb_BaseProduct_Click()
    'Call loadDataCampaign
    Call ShowComboBox
    Call loadDataUser
    Call tampil_reason
End Sub
Public Sub ShowComboBox(Optional isiAwal As String)
    sWhere = ""
    STRSQL = "select * from tbluser WHERE 1=1"
    
    Select Case UCase(sLevel)
        Case UCase("Supervisor")
            sWhere = " and tbluser_userid = '" & sUser & "' AND tbluser_kdstatus ='1' AND f_cuti = 0"
        Case UCase("Manager")
            sWhere = " and tbluser_userid in (select tbluser_userid from tbluser where tbluser_mgrcode = '" & sUser & "') AND tbluser_kdstatus ='1' AND f_cuti = 0 AND tbluser_kdlevel = '2'"
        Case UCase("Branch Manager")
            sWhere = " and tbluser_userid in (select tbluser_userid from tbluser where bm_code = '" & sUser & "') AND tbluser_kdstatus ='1' AND f_cuti = 0 AND tbluser_kdlevel = '5'"
    End Select
    If Cmb_BaseProduct.Text = "MORTGAGE" Then
        sWhere = sWhere + vbCrLf + " and F_CP = '1' AND F_KPM='0'"
    ElseIf Cmb_BaseProduct.Text = "KPM" Then
        sWhere = sWhere + vbCrLf + " and F_CP = '1' AND F_KPM='1'"
    ElseIf Cmb_BaseProduct.Text = "TOPUP PL" Then
        sWhere = sWhere + vbCrLf + " and F_CP = '0'"
    End If
    STRSQL = STRSQL & sWhere
    
    Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
    Cmb_ShowUser.Clear
    Cmb_ShowUser.Text = isiAwal
    If isiAwal = "" Then
        Cmb_ShowUser.Text = sUserid
        If UCase(sLevel) = "MANAGER" Then
            Cmb_ShowUser.Text = "[ALL]"
        ElseIf UCase(sLevel) = "BRANCH MANAGER" Then
            Cmb_ShowUser.Text = "[AGENT]"
        End If
    End If
    
    If UCase(sLevel) = "MANAGER" Then
        Cmb_ShowUser.AddItem "[ALL]"
        If MDIForm1.txt_f_cp.Text = "0" Then
            Cmb_ShowUser.AddItem "[ALL AGENT]"
        End If
    ElseIf UCase(sLevel) = "BRANCH MANAGER" Then
        Cmb_ShowUser.AddItem "[ALL]"
        Cmb_ShowUser.AddItem "[MANAGER]"
        Cmb_ShowUser.AddItem "[SPV]"
        Cmb_ShowUser.AddItem "[AGENT]"
    End If
    While Not rs.EOF
        'Cmb_ShowUser.AddItem cnull(rs!tbluser_userid)
        Cmb_ShowUser.AddItem cnull(rs!tbluser_userid) & " - " & cnull(rs!tbluser_name)
        rs.MoveNext
    Wend
    
    Set rs = Nothing
End Sub
Private Sub Cmb_ShowUser_Click()
    Call loadDataUser
    Call tampil_reason
    isiAwalShowUser = Cmb_ShowUser.Text
End Sub

Private Sub cmbCriteria_DropDown()
    Call getKriteria
End Sub

Private Sub getKriteria()
    Dim rs As ADODB.Recordset
    Dim STRSQL, MWHERE, strCamp As String
    
    strCamp = getcampaign
    
    If strCamp <> Empty Then
        STRSQL = " select distinct kriteria from mgm "
        MWHERE = " where 1=1"
        MWHERE = MWHERE + " and campaign_code in (" + strCamp + ")  "
    Else
        STRSQL = " select kriteria from tbl_kriteria  "
        MWHERE = " where 1=1"
    End If
    
    STRSQL = STRSQL + MWHERE + " order by kriteria"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
    cmbCriteria.Clear
    While Not rs.EOF
        cmbCriteria.AddItem cnull(rs!kriteria)
        rs.MoveNext
    Wend
    
    Set rs = Nothing

End Sub

Private Sub getBaseProduct()
    Dim rs As ADODB.Recordset
    Dim STRSQL, MWHERE As String
    
    
    STRSQL = " select keterangan from tblprogram"
    MWHERE = " where 1=1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
    Cmb_BaseProduct.Clear
    While Not rs.EOF
        Cmb_BaseProduct.AddItem cnull(rs!keterangan)
        rs.MoveNext
    Wend
    
    Set rs = Nothing
End Sub

Private Sub Cmd_Load_Click()
    Dim manager_boolean As Boolean
    Dim campaign_boolean As Boolean
    Dim i, jumlah_data, bagi_dua, ganjil, count_bagi, bagi_ke As Integer
    List_Proses_Data.ColumnHeaders.Clear
    List_Proses_Data.ColumnHeaders.Add 1, , "Campaign Code", 2500
    Cmd_Load.Enabled = False
    manager_boolean = False
    count_bagi = 0
    K = 1
    For i = 1 To List_Manager.ListItems.Count
        If List_Manager.ListItems(i).Checked = True Then
            List_Proses_Data.ColumnHeaders.Add K + 1, , List_Manager.ListItems(i).Text
            count_bagi = count_bagi + 1
            K = K + 1
            manager_boolean = True
        End If
    Next i
    If manager_boolean = False Or K > 3 Then
        Call msg("Manager Tidak Boleh Kosong Atau Tidak Boleh Dari 2")
        Exit Sub
    End If
    campaign_boolean = False
    bagi_ke = 1
    List_Proses_Data.ListItems.Clear
    
    For i = 1 To List_Campaign.ListItems.Count
        If List_Campaign.ListItems(i).Checked = True Then
            jumlah_data = List_Campaign.ListItems(i).SubItems(1)
            bagi_dua = Int(jumlah_data / count_bagi)
            ganjil = jumlah_data - (bagi_dua * count_bagi)
            
            
            Set listv = List_Proses_Data.ListItems.Add(1, , List_Campaign.ListItems(i).Text)
            For j = 1 To count_bagi
                If ganjil = 1 Then
                    If bagi_ke = j Then
                        listv.SubItems(j) = bagi_dua + 1
                        If bagi_ke = 2 Then
                            bagi_ke = 1
                        Else
                            bagi_ke = 2
                        End If
                        ganjil = 0
                    Else
                        listv.SubItems(j) = bagi_dua
                        ganjil = 1
                    End If
                Else
                    listv.SubItems(j) = bagi_dua
                End If
            Next j
            campaign_boolean = True
        End If
    Next i
    If campaign_boolean = False Then
        Call msg("Campaign Code Tidak Boleh Kosong")
        Exit Sub
    End If
    Cmd_Proses.Enabled = True
    Cmd_Load.Enabled = True
End Sub

Private Sub Cmd_proses_Click()
    Dim proses_data As Boolean
    Dim campaign_code, user, nama, jumlah_data, sQueryid As String
    Cmd_Proses.Enabled = False
    Cmd_Load.Enabled = False
    proses_data = False
    For i = 1 To List_Proses_Data.ListItems.Count
        If List_Proses_Data.ListItems(1).Checked = True Then
            proses_data = True
            sWhere = ""
            sWhere = cTable.CheckWhere(sWhere, "(agent is null or agent='' or agent='" & MDIForm1.txtUserName.Text & "')")
            sWhere = cTable.CheckWhere(sWhere, "statuscall='New Data'")
            
            campaign_code = List_Proses_Data.ListItems(1).Text
            sWhere = cTable.CheckWhere(sWhere, "campaign_code = '" & campaign_code & "'")
            
            For j = 1 To List_Proses_Data.ColumnHeaders.Count - 1
                user = List_Proses_Data.ColumnHeaders(j + 1).Text
                nama = get_name_user(user)
                jumlah_data = List_Proses_Data.ListItems(1).SubItems(j)
                sQueryid = "SELECT id FROM mgm " & sWhere & " ORDER BY id LIMIT " & jumlah_data
                DoEvents
                'Call CreateInsert_Waterfall_Hst(sQueryid, user, nama, MDIForm1.txtlevel.Text)
                
                sQueryUpdate = " UPDATE mgm " & _
                               " SET agent='" + user + "', nmagent= '" + nama + "', tgl_distribusi =date(now()) " & _
                               " WHERE id in( " & sQueryid & " ) and statuscall <> 'Agree'  "
                DoEvents
                M_OBJCONN.Execute sQueryUpdate
                 
                sQueryInsert = " INSERT INTO tbllogdistribusi(userid,nama,campaign_code,jmldata,sendby) values " & _
                               " ('" & user & "'," & _
                               " '" & nama & "'," & _
                               " '" & campaign_code & "'," & _
                               CStr(Val(jumlah_data)) & "," & _
                               " '" + MDIForm1.txtUserName.Text + "')"
                DoEvents
                M_OBJCONN.Execute sQueryInsert
            Next j
            List_Proses_Data.ListItems.Remove (1)
        End If
    Next i
    Cmd_Load.Enabled = True
    If proses_data = False Then
        msg ("Silahkan Pilih List Proses Data")
        Cmd_Proses.Enabled = True
        Exit Sub
    End If
    msg ("Distribusi Otomatis Sukses")
    CheckAll_Campaign.Value = vbUnchecked
    CheckAll_ProsesData.Value = vbUnchecked
    Cmb_BaseProduct_Click
    List_Proses_Data.ListItems.Clear
    
End Sub
Function get_name_user(sUserid)
    Call UnConnectRs(rs)
    sQuerySelect = "SELECT tbluser_name FROM tbluser WHERE tbluser_userid='" & sUserid & "'"
    rs.open sQuerySelect
    If rs.EOF Then
        get_name_user = ""
        Exit Function
    End If
    get_name_user = cnull(rs!tbluser_name)
End Function

Private Sub CreateInsert_Waterfall_Hst(sQueryid As String, ByVal agent As String, ByVal nmAgent As String, Level As String)
    Dim sBulan      As String
    Dim sYear       As String
    Dim sTableName  As String
    
    sBulan = Format(FungsiWaktuServer, "mmmm")
    sYear = Format(FungsiWaktuServer, "yyyy")
    sTableName = "tbl_mgm_hst_" & sBulan & "_" & sYear
    On Error GoTo InsertTable
    'Call CreateTableMgmHst
InsertTable:
    Level = UCase(Level)
    If Level = "AGENT" Then
        sQuerySelect = " select id as id_cust,'New Data'::text,'New Data'::text,'" & agent & "'::text,'" & nmAgent & "'::text,tbluser_groupspvcode,tbluser_ketgroupspv,now(),campaign_code,campaign_name from mgm a,tbluser b where a.agent=b.tbluser_userid " & _
                       " and id in ( " & sQueryid & " ) and statuscall <> 'Agree'  "
    Else
        sQuerySelect = " select id as id_cust,'New Data'::text,'New Data'::text,'" & agent & "'::text,'" & nmAgent & "'::text,'" & agent & "'::text,'" & nmAgent & "'::text,now(),campaign_code,campaign_name from mgm where  " & _
                       " id in ( " & sQueryid & " ) and statuscall <> 'Agree'  "
    End If
                   
    sQueryInsert = " INSERT INTO " & sTableName & "(" & _
                   " id_cust,statuscall,reasoncall,agent,nmagent,kdspv,nmspv,tglcall,campaign_code,campaign_name " & _
                   " ) " & sQuerySelect
    M_OBJCONN.Execute sQueryInsert
End Sub

Private Sub cmdcancel_Click()
    frmConfirmExeption.Visible = False
End Sub

Private Sub cmdCancel2_Click()
    frmUserAssign.Visible = False
End Sub

Private Sub cmdException_Click()
    Dim list As ListItem

    With List_Exception
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No.", 5 * txt
        .ColumnHeaders.Add , , "Userid", 10 * txt
        .ColumnHeaders.Add , , "Nama", 0 * txt
        .ColumnHeaders.Add , , "Reason", 30 * txt
    End With
    
    If List_User.ListItems.Count = 0 Then
        Exit Sub
    End If
    For i = 1 To List_Exception.ListItems.Count
        If List_Exception.ListItems(i).SubItems(1) = List_User.SelectedItem.SubItems(1) Then
            MsgBox "Data ini Telah Dipilih", vbInformation, "Informasi"
            Exit Sub
        End If
    Next i
    txtUserException.Text = List_User.SelectedItem.SubItems(1)
    sExcepTrans = True
    frmConfirmExeption.Visible = True
    txtReasonException.SetFocus
End Sub

Private Sub cmdOKException_Click()
    
    If sExcepTrans = True Then
        Set LST = List_Exception.ListItems.Add(, , List_User.SelectedItem.Text)
        LST.SubItems(1) = List_User.SelectedItem.SubItems(1)
        LST.SubItems(2) = List_User.SelectedItem.SubItems(2)
        LST.SubItems(3) = txtReasonException.Text
        List_User.ListItems.Remove (List_User.SelectedItem.Index)
        Call itungJmlUser
        If txtTotaltUser <> "0" Then
            Call bagiRata
        End If
        Call insertLogException
        Call resetException
        Call ShowComboBox(isiAwalShowUser)
        Call loadDataUser
        Call tampil_reason
    Else
        List_Exception.SelectedItem.SubItems(3) = txtReasonException.Text
        Call resetException
    End If
End Sub

Private Sub cmdProsesDistribusi_Click()
    If txtTotaltUser.Text = "0" Then
        MsgBox "Tidak ada User", vbCritical + vbOKOnly, "TINS"
        Exit Sub
    End If
    If Val(txtTotalGetData.Text) > Val(txtTotalDtaCek.Text) Then
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            If MsgBox("Jumlah data(" + txtTotalDtaCek.Text + ") kurang dari jumlah data(" + txtTotalGetData.Text + ") yang akan di distribusi,yakin akan di proses?", vbQuestion + vbYesNo, "TINS") = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "Jumlah data(" + txtTotalDtaCek.Text + ") kurang dari jumlah data(" + txtTotalGetData.Text + ") yang akan di distribusi", vbCritical + vbOKOnly, "TINS"
            Exit Sub
        End If
    End If
    
    If Val(txtTotalGetData.Text) < Val(Replace(tdbDataToDistribute.Text, ",", "")) Then
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            If MsgBox("Jumlah data(" + txtTotalGetData.Text + ") kurang dari jumlah data(" + tdbDataToDistribute.Text + ") yang akan di distribusi,yakin akan di proses?", vbQuestion + vbYesNo, "TINS") = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "Jumlah data(" + txtTotalGetData.Text + ") kurang dari jumlah data(" + tdbDataToDistribute.Text + ") yang akan di distribusi", vbCritical + vbOKOnly, "TINS"
            Exit Sub
        End If
    End If
    
    If Val(txtTotalGetData.Text) > Val(Replace(tdbDataToDistribute.Text, ",", "")) Then
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            If MsgBox("Jumlah data(" + txtTotalGetData.Text + ") lebih dari jumlah data(" + tdbDataToDistribute.Text + ") yang akan di distribusi,yakin akan di proses?", vbQuestion + vbYesNo, "TINS") = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "Jumlah data(" + txtTotalGetData.Text + ") lebih dari jumlah data(" + tdbDataToDistribute.Text + ") yang akan di distribusi", vbCritical + vbOKOnly, "TINS"
            Exit Sub
        End If
    End If

    cmdProsesDistribusi.Enabled = False
    Call Prosesdistribusi
    cmdProsesDistribusi.Enabled = True
End Sub

Private Sub insertLogException()
    Dim CMDSQL, sUserException, sExceptionReason As String
                        
        CMDSQL = " insert into tbl_log_distribute_exception (userid,nama,exception_reason,userid_execute) "
        CMDSQL = CMDSQL + " values ('" + txtUserException.Text + "','" + List_Exception.SelectedItem.SubItems(2) + "','" + txtReasonException.Text + "','" + sUser + "');" + vbCrLf
        M_OBJCONN.Execute CMDSQL
        
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            CMDSQL = "update tbluser set f_cuti_bm='1', exception_reason = '" + txtReasonException.Text + "' where tbluser_userid='" + txtUserException.Text + "'"
        Else
            CMDSQL = "update tbluser set f_cuti='1', exception_reason = '" + txtReasonException.Text + "' where tbluser_userid='" + txtUserException.Text + "'"
        End If
        M_OBJCONN.Execute CMDSQL
End Sub
Private Sub insertLogDistribusi()
    Dim CMDSQL, sUserDistrib, sExceptionReason As String
    On Error GoTo ERRORA
    
    CMDSQL = Empty
    If List_User.ListItems.Count > 0 Then
        For i = 1 To List_User.ListItems.Count
            sUserDistrib = List_User.ListItems(i).SubItems(1)
            sNamaDistrib = List_User.ListItems(i).SubItems(2)
            sCampaignCde = getcampaign("1")
            sJmlDataDist = List_User.ListItems(i).SubItems(3)
                        
            CMDSQL = CMDSQL + " insert into tbllogdistribusi (userid,nama,campaign_code,jmldata,sendby) "
            CMDSQL = CMDSQL + " values ('" + sUserDistrib + "','" + sNamaDistrib + "','" + sCampaignCde + "','" + sJmlDataDist + "','" + sUser + "');" + vbCrLf
        Next i
        M_OBJCONN.Execute CMDSQL
    End If

Exit Sub
ERRORA:
    MsgBox "Error Loging!:" + err.Description, vbInformation + vbOKOnly, "TINS"
End Sub

Private Sub itungJmlUser()
    txtTotaltUser.Text = CStr(List_User.ListItems.Count)
End Sub

Private Function getcampaign(Optional sCode As String) As String
    Dim strCampaign As String
    Dim sC As String
    
    sC = IIf(sCode = "1", "", "'")
    For i = 1 To List_Campaign.ListItems.Count
        If List_Campaign.ListItems(i).Checked = True Then
            If strCampaign = Empty Then
                strCampaign = strCampaign + sC + List_Campaign.ListItems(i).SubItems(1) + sC
            Else
                strCampaign = strCampaign + "," + sC + List_Campaign.ListItems(i).SubItems(1) + sC + ""
            End If
        End If
    Next i
    
    getcampaign = strCampaign
End Function

Private Sub Prosesdistribusi()
On Error GoTo ke
    Dim rs As ADODB.Recordset
    Dim Ij, adk As Integer
    Dim STRSQL, strCampaign, sUserPil, sLmit, SELECT_LOG, CRT_LOG As String
            
    If List_Campaign.ListItems.Count = 0 Then
        MsgBox "Pilih Campaign Terlebih dahulu", vbOKOnly + vbCritical, "TINS"
        Exit Sub
    End If
                
    If Val(txtTotalGetData.Text) = 0 Then
        MsgBox "Isi jumlah data yang akan didistribusi", vbCritical + vbOKOnly, "TINS"
        Exit Sub
    End If
            
    strCampaign = getcampaign
    If strCampaign = Empty Then
        MsgBox "Pilih Campaign Terlebih Dahulu", vbCritical + vbOKOnly, "TINS"
        Exit Sub
    End If
    
    adk = 0
    For Ij = 1 To List_Campaign.ListItems.Count
        If List_Campaign.ListItems(Ij).Checked = True Then
            adk = adk + 1
        End If
    Next Ij
    
    If adk > 1 Then
        msg ("untuk sementara pembagian data hanya bisa per satu campaign")
        Exit Sub
    End If
            
    If MsgBox("Anda yakin ingin melakukan distribusi?", vbQuestion + vbYesNo, "TINS") = vbNo Then
        Exit Sub
    End If
            
    For i = 1 To List_User.ListItems.Count
        sUserPil = List_User.ListItems(i).SubItems(1)
        sLmit = CStr(Val(List_User.ListItems(i).SubItems(3)))

        If Val(sLmit) > 0 Then
            SELECT_LOG = " select id from mgm where 1=1 "
            SELECT_LOG = SELECT_LOG + vbCrLf + " and campaign_code in (" + strCampaign + ") "
            SELECT_LOG = SELECT_LOG + vbCrLf + " and agent = '" + sUser + "' "
            SELECT_LOG = SELECT_LOG + vbCrLf + " and coalesce(f_block,0) = '0' "
            SELECT_LOG = SELECT_LOG + vbCrLf + " and campaign_code not in (select distinct tbldatasource_campaign_code from tbldatasource where tbldatasource_kdstatus = '0')"
            '--block campaign FIRAS
            If MDIForm1.Txtlevel.Text = "Supervisor" Then
                SELECT_LOG = SELECT_LOG & " and campaign_code not in (select distinct campaign_code from tblblockcampaign_tmp where agent ='" & MDIForm1.txtUserName.Text & "' and status = '5' and exec_user in (select tbluser_mgrcode from tbluser where tbluser_userid = '" & MDIForm1.txtUserName.Text & "'))"
            End If
            '--
            If cmbCriteria.Text <> Empty Then
                SELECT_LOG = SELECT_LOG + vbCrLf + " and kriteria = '" + cmbCriteria.Text + "' "
            End If
            
            If Cmb_BaseProduct.Text <> Empty Then
                SELECT_LOG = SELECT_LOG + vbCrLf + " and base_product = '" + Cmb_BaseProduct.Text + "' "
            End If
            
            SELECT_LOG = SELECT_LOG + vbCrLf + " order by row_number() over(partition by campaign_code) "
            SELECT_LOG = SELECT_LOG + vbCrLf + " limit " + sLmit
            
            CMDSQL = "update mgm set agent = '" + sUserPil + "', nmagent = '" & List_User.ListItems(i).SubItems(2) & "',tgl_distribusi= now() where id in (" + SELECT_LOG + " ); "
            '===================================================================
            CRT_LOG = "insert into tbl_log_distribute_auto (to_userid,from_userid,id_mgm) "
            CRT_LOG = CRT_LOG + vbCrLf + " select '" + sUserPil + "','" + sUser + "',id from (" + SELECT_LOG + ") as a;"
            
            M_OBJCONN.Execute CMDSQL + CRT_LOG
        End If
    Next i
    
    
    Call insertLogDistribusi
    MsgBox "Proses Distribusi Done", vbInformation + vbOKOnly, "TINS"
    
    loadDataCampaign
    loadDataUser
Exit Sub
ke:
MsgBox err.Description, vbInformation + vbOKOnly, "TINS"

End Sub

Private Sub Command1_Click()
    If List_Exception.ListItems.Count <> 0 Then
        Set lList = List_User.ListItems.Add(, , List_Exception.SelectedItem.Text)
        lList.SubItems(1) = List_Exception.SelectedItem.SubItems(1)
        lList.SubItems(2) = List_Exception.SelectedItem.SubItems(2)
        List_Exception.ListItems.Remove List_Exception.SelectedItem.Index
        
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            CMDSQL = "update tbluser set f_cuti_BM='0', exception_reason = '' where tbluser_userid='" + lList.SubItems(1) + "'        "
        Else
            CMDSQL = "update tbluser set f_cuti='0', exception_reason = '' where tbluser_userid='" + lList.SubItems(1) + "'        "
        End If
        M_OBJCONN.Execute CMDSQL
        
        CMDSQL = "update tbl_log_distribute_exception set flag='1' where userid='" + lList.SubItems(1) + "'        "
        M_OBJCONN.Execute CMDSQL
        Call ShowComboBox(isiAwalShowUser)
        Call loadDataUser
        Call tampil_reason
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 14175
    Me.Height = 10935
    sExcepTrans = False
    sUser = MDIForm1.txtUserName.Text
    sLevel = MDIForm1.Txtlevel.Text
    'Call getBaseProduct
    Cmb_BaseProduct.Clear
    Cmb_BaseProduct.AddItem "CC REGULER"
    Cmb_BaseProduct.AddItem "TOPUP REGULER"
    Cmb_BaseProduct.AddItem "PL EXPRESS"
    Cmb_BaseProduct.AddItem "TRANSFER TUNAI"
    Cmb_BaseProduct.AddItem "LUNAS"
    
    If MDIForm1.Txtlevel.Text = "Branch Manager" Then

        Cmb_BaseProduct.Enabled = True
        Cmb_ShowUser.AddItem "[MANAGER]"
        Cmb_ShowUser.AddItem "[SUPERVISOR]"
        cmbflag.Visible = True
        Label1(2).Visible = True
        Label18.Visible = True
        tdbDatarata.Visible = True
        Label17.Visible = True
        txttotalspv.Visible = True
        'Call getFlag
    End If
    Call ShowComboBox
    Call loadDataUser
    Call tampil_reason
End Sub
Private Sub getFlag()
    Dim rs As ADODB.Recordset
    Dim STRSQL, MWHERE, strCamp As String
    
    
    If Cmb_BaseProduct.Text = "MORTGAGE" Then
        STRSQL = " select distinct f_param from tbluser where f_cp='1' and f_kpm='0'"
    ElseIf Cmb_BaseProduct.Text = "KPM" Then
        STRSQL = " select distinct f_param from tbluser where f_cp='1' and f_kpm='1'"
    ElseIf Cmb_BaseProduct.Text = "TOPUP PL" Then
        STRSQL = " select distinct f_param from tbluser where f_cp='0'"
    End If
    
    STRSQL = STRSQL + MWHERE + " order by f_param"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
    cmbflag.Clear
    While Not rs.EOF
        cmbflag.AddItem cnull(rs!f_param)
        rs.MoveNext
    Wend
    
    Set rs = Nothing

End Sub
Private Sub Form_Unload(cancel As Integer)
    Call UnConnectRs(rs)
End Sub
Private Sub cmdLoadDataCampaign_Click()
    cmdLoadDataCampaign.Enabled = False
    Call loadDataCampaign
    Call loadDataUser
    cmdLoadDataCampaign.Enabled = True
End Sub
Private Sub loadDataCampaign()
    Dim rs As ADODB.Recordset
    Dim list As ListItem
    Dim STRSQL, sUserName, sField As String
    Dim MHWERE As String
    Dim sJ As Double
    
    
    STRSQL = "select"
    STRSQL = STRSQL + vbCrLf + " campaign_code"
    STRSQL = STRSQL + vbCrLf + " ,sum(case when agent='" + sUser + "' and m.flag_recycle = 0 and m.bucket = 'N' then 1 else 0 end) as remain"
    STRSQL = STRSQL + vbCrLf + " ,count(m.id) as total_data"
    STRSQL = STRSQL + vbCrLf + " ,sum(case when agent='" + sUser + "' and m.flag_recycle = 1 and m.bucket = 'R' then 1 else 0 end) as recycle"
    STRSQL = STRSQL + vbCrLf + " from mgm m"
    STRSQL = STRSQL + vbCrLf + " left join tbldatasource d on (d.tbldatasource_campaign_code=m.campaign_code)"
    
    MWHERE = vbCrLf + " Where 1 = 1"
    MWHERE = MWHERE + vbCrLf + " and date(tbldatasource_tglexpired)>date(now())"
    MWHERE = MWHERE + vbCrLf + " and (date(startdate)<=date(now()) or startdate is null)"
    MWHERE = MWHERE + vbCrLf + " and coalesce(m.f_block,0) = '0'"
    MWHERE = MWHERE + vbCrLf + " and d.tbldatasource_kdstatus = '1'"
    
    If MDIForm1.Txtlevel.Text = "Manager" Then
        MWHERE = MWHERE + vbCrLf + " and (d.officer_id = '" & MDIForm1.txtUserName.Text & "' )"
    End If
    '--block campaign FIRAS
    If MDIForm1.Txtlevel.Text = "Supervisor" Then
        MWHERE = MWHERE & " and m.campaign_code not in (select distinct campaign_code from tblblockcampaign_tmp where agent ='" & MDIForm1.txtUserName.Text & "' and status = '5' and exec_user in (select tbluser_mgrcode from tbluser where tbluser_userid = '" & MDIForm1.txtUserName.Text & "')) "
    End If
    '--
    sField = ""
    If sLevel = "Manager" Then
        sField = "tbluser_mgrcode"
    ElseIf sLevel = "Supervisor" Then
        sField = "tbluser_groupspvcode"
    Else
        sField = "tbluser_userid"
    End If
    
    MWHERE = MWHERE + vbCrLf + " and (agent ='" + sUser + "' or agent in (select tbluser_userid from tbluser where " + sField + "='" + sUser + "'))"
    
    If Cmb_BaseProduct.Text <> Empty Then
        MWHERE = MWHERE + vbCrLf + " and base_product = '" + Cmb_BaseProduct.Text + "'"
    End If
    
    If cmbCriteria.Text <> Empty Then
        MWHERE = MWHERE + vbCrLf + " and kriteria = '" + cmbCriteria.Text + "'"
    End If
    
    
    If Txt_Search.Text <> Empty Then
        MWHERE = MWHERE + vbCrLf + " and tbldatasource_campaign_code  like '%" + Txt_Search.Text + "%'"
    End If
    
    STRSQL = STRSQL + MWHERE + vbCrLf + " group by campaign_code order by campaign_code"
    
    'Call UnConnectRs(rs)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    txtJmlCampaign.Text = CStr(Val(rs.RecordCount))
    
    If rs.RecordCount = 0 Then
        List_Campaign.ListItems.Clear
        MsgBox "Data Not Found", vbInformation + vbOKOnly, "TINS"
        Exit Sub
    End If
            
    With List_Campaign
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No.", 5 * txt
        .ColumnHeaders.Add , , "Campaign Code", 55 * txt
        .ColumnHeaders.Add , , "Remain", 10 * txt
        .ColumnHeaders.Add , , "Total Data", 10 * txt
        .ColumnHeaders.Add , , "Recycle", 10 * txt
        
        .ListItems.Clear
        
        sJ = 0
        List_Campaign.ListItems.Clear
        While Not rs.EOF
            Set list = .ListItems.Add(, , rs.Bookmark)
            list.SubItems(1) = cnull(rs!campaign_code)
            list.SubItems(2) = cnull(rs!Remain)
            list.SubItems(3) = cnull(rs!total_data)
            list.SubItems(4) = cnull(rs!recycle)
            sJ = sJ + Val(cnull(rs!Remain))
            rs.MoveNext
        Wend
        txtTotal.Text = CStr(Val(sJ))
        'tdbDataToDistribute.Value = Val(sJ)
    End With
    Call itungJmlCek
    Set rs = Nothing
End Sub
Private Sub cmdLoadUser_Click()
    cmdLoadUser.Enabled = False
    Call loadDataUser
    Call tampil_reason
    cmdLoadUser.Enabled = True
End Sub

Private Sub loadDataUser()
    Dim rs As ADODB.Recordset
    Dim list As ListItem
    Dim STRSQL, strCampaign As String
    Dim MHWERE As String
    Dim getCode As String
    Dim getName As String
    
    '=================================================
    getCode = ""
    getName = ""
    txttotalspv.Text = ""
    intvrl = InStr(1, Cmb_ShowUser.Text, " - ", vbTextCompare)
    If intvrl <> 0 Then
       ArrayString = Split(Cmb_ShowUser.Text, " - ", 2, vbTextCompare)
       getCode = ArrayString(0)
       getName = ArrayString(1)
    End If
    '=================================================
    
    strCampaign = getcampaign
    If MDIForm1.Txtlevel.Text = "Branch Manager" Then
        STRSQL = " select f_param,"
        STRSQL = STRSQL + vbCrLf + "  tbluser_userid As USERID"
        STRSQL = STRSQL + vbCrLf + "  ,tbluser_name as nama"
        STRSQL = STRSQL + vbCrLf + "  ,u.tbluser_groupspvcode as SPV"
        STRSQL = STRSQL + vbCrLf + "  ,total_agent"
        STRSQL = STRSQL + vbCrLf + "  ,total_spv"
    Else
        STRSQL = " select f_param,"
        STRSQL = STRSQL + vbCrLf + "  tbluser_userid As USERID"
        STRSQL = STRSQL + vbCrLf + "  ,tbluser_name as nama"
    End If
    
    If strCampaign <> Empty And chkShowAssign.Value = 1 Then
        STRSQL = STRSQL + vbCrLf + "  ,coalesce(a.jml_distributed,'0') as jml_distributed"
    End If
    
    If MDIForm1.Txtlevel.Text = "Branch Manager" Then
        STRSQL = STRSQL + vbCrLf + " from tbluser u"
        STRSQL = STRSQL + vbCrLf + " LEFT JOIN (select count(distinct tbluser_userid) as total_agent,tbluser_groupspvcode from tbluser where tbluser_kdlevel = '1'"
        STRSQL = STRSQL + vbCrLf + " and tbluser_kdstatus = '1' group by tbluser_groupspvcode) t on (u.tbluser_groupspvcode=t.tbluser_groupspvcode)"
        STRSQL = STRSQL + vbCrLf + " LEFT JOIN (select count(distinct tbluser_groupspvcode) as total_spv,tbluser_mgrcode from tbluser where tbluser_kdlevel = '2'"
        STRSQL = STRSQL + vbCrLf + " and tbluser_kdstatus = '1' group by tbluser_mgrcode) v on (u.tbluser_mgrcode=v.tbluser_mgrcode)"
        'strsql = strsql + vbCrLf + " inner join (select distinct userid from tblloguseradm where date(tgl)=date(now()) and operation='Success Login') w on (u.tbluser_userid=w.userid)"
    Else
        STRSQL = STRSQL + vbCrLf + " from tbluser u"
    End If
    
    If strCampaign <> Empty And chkShowAssign.Value = 1 Then
        STRSQL = STRSQL + vbCrLf + " left join ("
        STRSQL = STRSQL + vbCrLf + "    select"
        
        If sLevel = "Manager" Then
            STRSQL = STRSQL + vbCrLf + "    case when coalesce(tbluser_groupspvcode,'')='' then agent else tbluser_groupspvcode end as user_group"
        ElseIf sLevel = "Supervisor" Or sLevel = "Branch Manager" Then
            STRSQL = STRSQL + vbCrLf + "    agent as user_group"
        End If
        
        STRSQL = STRSQL + vbCrLf + "    ,count(id) as jml_distributed"
        STRSQL = STRSQL + vbCrLf + "    from mgm m"
        STRSQL = STRSQL + vbCrLf + "    left join tbluser u on (u.tbluser_userid=m.agent)"
        STRSQL = STRSQL + vbCrLf + "    Where 1 = 1"
        STRSQL = STRSQL + vbCrLf + "    and campaign_code in (" + strCampaign + ") "
        STRSQL = STRSQL + vbCrLf + "group by user_group"
        STRSQL = STRSQL + vbCrLf + ")as a on (a.user_group=u.tbluser_userid)"
    End If
    
    If sLevel = "Manager" Then
        If Cmb_ShowUser.Text = "[ALL]" Then
            MWHERE = " Where f_cuti='0'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_mgrcode ='" + sUser + "'"
            If MDIForm1.txt_f_cp.Text = "0" Then
                MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '2'"
            Else
                MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '1'"
                If Cmb_BaseProduct.Text = "KPM" Then
                    MWHERE = MWHERE + vbCrLf + " and F_KPM = '1'"
                Else
                     MWHERE = MWHERE + vbCrLf + " and F_KPM = '0'"
                End If
            End If
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
        ElseIf Cmb_ShowUser.Text = "[ALL AGENT]" Then
            MWHERE = " Where f_cuti in('0','1')"
            MWHERE = MWHERE + vbCrLf + " and tbluser_mgrcode ='" + sUser + "'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '1'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
                If Cmb_BaseProduct.Text = "KPM" Then
                    MWHERE = MWHERE + vbCrLf + " and F_KPM = '1'"
                Else
                     MWHERE = MWHERE + vbCrLf + " and F_KPM = '0'"
                End If
        Else
            If MDIForm1.txt_f_cp.Text = "0" Then
                MWHERE = " Where f_cuti='0'"
                MWHERE = MWHERE + vbCrLf + " and tbluser_groupspvcode ='" + Cmb_ShowUser.Text + "'"
                MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '1'"
                MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
            Else
                MWHERE = " Where f_cuti='0'"
                MWHERE = MWHERE + vbCrLf + " and tbluser_mgrcode ='" + Cmb_ShowUser.Text + "'"
                MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '1'"
                MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
                If Cmb_BaseProduct.Text = "KPM" Then
                    MWHERE = MWHERE + vbCrLf + " and F_KPM = '1'"
                Else
                     MWHERE = MWHERE + vbCrLf + " and F_KPM = '0'"
                End If
            End If
        End If
    ElseIf sLevel = "Supervisor" Then
        MWHERE = " Where f_cuti='0'"
        MWHERE = MWHERE + vbCrLf + " and tbluser_groupspvcode ='" + sUser + "'"
        MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '1'"
        MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
    ElseIf sLevel = "Branch Manager" Then
        MWHERE = " Where COALESCE(tbluser_userid,'')<>'' AND f_cuti_BM='0'"
        
        If Cmb_ShowUser.Text = "[ALL]" Then
            MWHERE = MWHERE + vbCrLf + " and bm_code ='" + sUser + "'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel in('5','2')"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
            If Cmb_BaseProduct.Text = "KPM" Or Cmb_BaseProduct.Text = "MORTGAGE" Then
                MWHERE = MWHERE + vbCrLf + " AND F_CP='1'"
            Else
                 MWHERE = MWHERE + vbCrLf + " AND F_CP = '0'"
            End If
        ElseIf Cmb_ShowUser.Text = "[MANAGER]" Then
            MWHERE = MWHERE + vbCrLf + " and bm_code ='" + sUser + "'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '5'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
            If Cmb_BaseProduct.Text = "KPM" Or Cmb_BaseProduct.Text = "MORTGAGE" Then
                MWHERE = MWHERE + vbCrLf + " AND F_CP='1'"
            Else
                 MWHERE = MWHERE + vbCrLf + " AND F_CP = '0'"
            End If
        ElseIf Cmb_ShowUser.Text = "[SPV]" Then
            MWHERE = MWHERE + vbCrLf + " and bm_code ='" + sUser + "'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '2'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
            If Cmb_BaseProduct.Text = "KPM" Or Cmb_BaseProduct.Text = "MORTGAGE" Then
                MWHERE = MWHERE + vbCrLf + " AND F_CP='1'"
            Else
                 MWHERE = MWHERE + vbCrLf + " AND F_CP = '0'"
            End If
        ElseIf Cmb_ShowUser.Text = "[AGENT]" Then
            MWHERE = MWHERE + vbCrLf + " and bm_code ='" + sUser + "'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '1'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
            If Cmb_BaseProduct.Text = "KPM" Then
                MWHERE = MWHERE + vbCrLf + " and F_KPM = '1' AND F_CP='1'"
            ElseIf Cmb_BaseProduct.Text = "MORTGAGE" Then
                MWHERE = MWHERE + vbCrLf + " and F_KPM = '0' AND F_CP='1'"
            Else
                 MWHERE = MWHERE + vbCrLf + " and F_KPM = '0' AND F_CP = '0'"
            End If
        ElseIf Mid(Cmb_ShowUser.Text, 1, 1) = "T" Or Mid(Cmb_ShowUser.Text, 1, 1) = "O" Then
            MWHERE = MWHERE + vbCrLf + " and tbluser_mgrcode ='" + getCode + "'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '2'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
        ElseIf Mid(Cmb_ShowUser.Text, 1, 1) = "S" Then
            MWHERE = MWHERE + vbCrLf + " and tbluser_GROUPSPVcode ='" + getCode + "'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdlevel = '1'"
            MWHERE = MWHERE + vbCrLf + " and tbluser_kdstatus = '1'"
        End If
        If cmbflag.Text <> "" Then
            MWHERE = MWHERE + vbCrLf + " and f_param = '" + cmbflag.Text + "'"
        End If
    End If
    '=================================================
    
    STRSQL = STRSQL + MWHERE + " order by random()"
    
    'Call UnConnectRs(rs)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    txtTotaltUser.Text = CStr(Val(rs.RecordCount))
    If rs.RecordCount <> 0 Then
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            txttotalspv.Text = cnull(rs!total_spv)
        End If
    
            
    With List_User
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No.", 5 * txt
        .ColumnHeaders.Add , , "userid", 10 * txt
        .ColumnHeaders.Add , , "name", 20 * txt
        .ColumnHeaders.Add , , "Data to Distribute", 10 * txt
        .ColumnHeaders.Add , , "Data Assign", 10 * txt
        .ColumnHeaders.Add , , "Flag", 10 * txt
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            .ColumnHeaders.Add , , "SPV", 10 * txt
            .ColumnHeaders.Add , , "Jml Agent", 10 * txt
            .ColumnHeaders.Add , , "Jml SPV", 10 * txt
        End If
        .ListItems.Clear
'        While Not rs.EOF
'            Set list = .ListItems.Add(, , rs.Bookmark)
'            list.SubItems(1) = cnull(rs!UserId)
'            list.SubItems(2) = cnull(rs!nama)
'            list.SubItems(3) = "0"
'            If strCampaign <> Empty And chkShowAssign.Value = 1 Then
'                list.SubItems(4) = cnull(rs!jml_distributed)
'            End If
'            rs.MoveNext
'        Wend
        
        While Not rs.EOF
            Set list = .ListItems.Add(, , rs.Bookmark)
            list.SubItems(1) = cnull(rs!UserId)
            If Mid(list.SubItems(1), 1, 5) = "TUPSP" Or Mid(list.SubItems(1), 1, 5) = "OMSPV" Then
                list.ListSubItems(1).ForeColor = vbRed
            Else
                list.ListSubItems(1).ForeColor = vbBlack
            End If
            list.SubItems(2) = cnull(rs!nama)
            list.SubItems(3) = "0"
            If strCampaign <> Empty And chkShowAssign.Value = 1 Then
                list.SubItems(4) = cnull(rs!jml_distributed)
            End If
            list.SubItems(5) = cnull(rs!f_param)
            If MDIForm1.Txtlevel.Text = "Branch Manager" Then
                list.SubItems(6) = cnull(rs!spv)
                list.SubItems(7) = cnull(rs!total_agent)
                list.SubItems(8) = cnull(rs!total_spv)
            End If
            rs.MoveNext
        Wend
        'txtJmlDatatoUser.Text = "0"
    End With
    End If

    Set rs = Nothing
    
End Sub
Private Sub loadDataUser_spv()
    Dim rs As ADODB.Recordset
    Dim list As ListItem
    Dim STRSQL, strCampaign As String
    Dim MHWERE As String
    Dim getCode As String
    Dim getName As String
    
    '=================================================
    If MDIForm1.Txtlevel.Text = "Branch Manager" Then
        STRSQL = " select distinct spv as spv1,jmlperspv from tbl_tmp_distribution_by_spv"
    End If
    '=================================================
    
    'Call UnConnectRs(rs)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If rs.RecordCount <> 0 Then
    With ListView1
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No.", 5 * txt
        .ColumnHeaders.Add , , "SPV", 10 * txt
        .ColumnHeaders.Add , , "Jumlah", 20 * txt
        .ListItems.Clear
        While Not rs.EOF
            Set list = .ListItems.Add(, , rs.Bookmark)
            list.SubItems(1) = cnull(rs!spv1)
            list.SubItems(2) = cnull(rs!jmlperspv)
            rs.MoveNext
        Wend
    End With
    End If

    Set rs = Nothing
    
End Sub
Private Sub tampil_reason()
Dim rs As ADODB.Recordset
Dim STRSQL As String
'    STRSQL = "select * from tbl_log_distribute_exception where userid_execute='" + MDIForm1.txtUserName.Text + "' and flag='0'"
    If MDIForm1.txt_f_cp.Text = "1" Then
        STRSQL = "select * from tbluser where tbluser_mgrcode='" + MDIForm1.txtUserName.Text + "' and tbluser_kdlevel = '1' and f_cuti='1' and tbluser_kdstatus = '1'"
    Else
        If UCase(MDIForm1.Txtlevel.Text) = "SUPERVISOR" Then
            STRSQL = "select * from tbluser where tbluser_groupspvcode='" + MDIForm1.txtUserName.Text + "' and tbluser_kdlevel = '1' and f_cuti='1' and tbluser_kdstatus = '1'"
        ElseIf UCase(MDIForm1.Txtlevel.Text) = "MANAGER" Then
            STRSQL = "select * from tbluser where tbluser_mgrcode='" + MDIForm1.txtUserName.Text + "' and tbluser_kdlevel = '2' and f_cuti='1' and tbluser_kdstatus = '1'"
        End If
    End If
    If UCase(MDIForm1.Txtlevel.Text) = "BRANCH MANAGER" Then
        If Cmb_BaseProduct.Text = "TOPUP PL" Then
            STRSQL = "select * from tbluser where BM_CODE='" + MDIForm1.txtUserName.Text + "'  and f_cuti_BM='1' and tbluser_kdstatus = '1' AND f_cp='0'"
        ElseIf Cmb_BaseProduct.Text = "MORTGAGE" Then
            STRSQL = "select * from tbluser where BM_CODE='" + MDIForm1.txtUserName.Text + "'  and f_cuti_BM='1' and tbluser_kdstatus = '1' and f_cp='1' AND f_kpm='0'"
        Else
            STRSQL = "select * from tbluser where BM_CODE='" + MDIForm1.txtUserName.Text + "' and f_cuti_BM='1' and tbluser_kdstatus = '1' and f_cp='1' AND f_kpm='1'"
        End If
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
     With List_Exception
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No.", 5 * txt
        .ColumnHeaders.Add , , "Userid", 10 * txt
        .ColumnHeaders.Add , , "Nama", 0 * txt
        .ColumnHeaders.Add , , "Reason", 30 * txt
    End With
    List_Exception.ListItems.Clear
    While Not rs.EOF
    no = no + 1
        Set list = List_Exception.ListItems.Add(, , no)
            list.SubItems(1) = cnull(rs!tbluser_userid)
            list.SubItems(2) = cnull(rs!tbluser_name)
            list.SubItems(3) = cnull(rs!exception_reason)
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub
Private Sub List_Campaign_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index = 1 Then 'ANGKA
        If List_Campaign.SortOrder = 0 Then
            Call SortColumn(List_Campaign, ColumnHeader.Index, sortDescending, sortNumeric)
        Else
            Call SortColumn(List_Campaign, ColumnHeader.Index, sortAscending, sortNumeric)
        End If
    Else
        If List_Campaign.SortOrder = 0 Then 'HURUF
            Call SortColumn(List_Campaign, ColumnHeader.Index, sortDescending, sortAlpha)
        Else
            Call SortColumn(List_Campaign, ColumnHeader.Index, sortAscending, sortAlpha)
        End If
    End If
End Sub

Private Sub List_Campaign_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call itungJmlCek
    Call bagiRata
End Sub

Private Sub itungJmlCek()
On Error GoTo a
    Dim sJ As Double
    Dim i As Integer
    
    sJ = 0
    For i = 1 To List_Campaign.ListItems.Count
        If List_Campaign.ListItems(i).Checked = True Then
            sJ = sJ + Val(List_Campaign.ListItems(i).SubItems(2))
        End If
    Next i
    txtTotalDtaCek.Text = CStr(Val(sJ))
    tdbDataToDistribute.Value = Val(sJ)
    If MDIForm1.Txtlevel.Text = "Branch Manager" Then
        If tdbDataToDistribute.Value <> 0 Then
            tdbDatarata.Value = tdbDataToDistribute.Value / txttotalspv.Text
        Else
            tdbDatarata.Value = 0
        End If
    End If
Exit Sub
a:
msg err.Description
End Sub

Private Sub List_Exception_DblClick()
    frmConfirmExeption.Visible = True
    txtUserException.Text = List_Exception.SelectedItem.SubItems(1)
    txtReasonException.Text = List_Exception.SelectedItem.SubItems(2)
    
End Sub

Private Sub List_User_Click()
        Call resetException
End Sub
Private Sub resetException()
    txtUserException.Text = Empty
    txtReasonException.Text = Empty
    frmConfirmExeption.Visible = False
    sExcepTrans = False
End Sub

Private Sub List_User_DblClick()
Dim rs As ADODB.Recordset
Dim STRSQL, MWHERE, strCampaign As String
    
strCampaign = getcampaign
If strCampaign <> Empty Then
    STRSQL = " select * from tbldatasource "
    MWHERE = " where 1=1"
    MWHERE = MWHERE + " and tbldatasource_campaign_code in (" + strCampaign + ")  "
    
    STRSQL = STRSQL + MWHERE
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
    If UCase(cnull(rs!distribusitype)) = "UNEVEN" Then
        If optUneven.Value = True Then
            txtUserAssign.Text = List_User.SelectedItem.SubItems(1)
            txtJmlAssign.Value = Val(List_User.SelectedItem.SubItems(3))
            frmUserAssign.Visible = True
            txtJmlAssign.SetFocus
        End If
    End If
    
    Set rs = Nothing
End If
End Sub
Private Sub cmdOk_Click()
    List_User.SelectedItem.SubItems(3) = CStr(Val(txtJmlAssign.Text))
    txtJmlAssign.Text = Empty
    frmUserAssign.Visible = False
    List_User.SetFocus
    Call totalGetData
End Sub

Private Sub List_User_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        List_User_DblClick
    End If
End Sub

Private Sub optEven_Click()
    If txtTotaltUser.Text = "0" Then
        MsgBox "Tidak ada User", vbCritical + vbOKOnly, "TINS"
        Exit Sub
    Else
        If MDIForm1.Txtlevel.Text = "Branch Manager" Then
            Call Bagi_rata_per_spv
            Call totalGetData
            Call loadDataUser_spv
        Else
            Call bagiRata
            Call totalGetData
        End If
    End If
End Sub

Sub Bagi_rata_per_spv()
On Error GoTo a
    M_OBJCONN.Execute "Delete from tbl_tmp_distribution_by_spv"
    For j = 1 To List_User.ListItems.Count
    
      XX = "  insert into  tbl_tmp_distribution_by_spv(agent,spv, jumlah_agent,jmlperspv) values ('" & List_User.ListItems(j).SubItems(1) & "','" & List_User.ListItems(j).SubItems(6) & "','" & List_User.ListItems(j).SubItems(7) & "','" & Val(tdbDatarata) & "')"
        M_OBJCONN.Execute XX
    Next j
    
    DoEvents
    M_OBJCONN.Execute "Select * from sp_bagirata_perspv ()"
    For i = 1 To List_User.ListItems.Count
        STRSQL = "select * from tbl_tmp_distribution_by_spv where agent='" + List_User.ListItems(i).SubItems(1) + "'"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
        If rs.RecordCount <> 0 Then
            List_User.ListItems(i).SubItems(3) = cnull(rs!jmldata)
        End If
    Next i
    Exit Sub
a:
    msg err.Description
End Sub
Private Sub optUneven_Click()
    Call totalGetData
End Sub

Private Sub bagiRata()
    Dim ibg As Integer
    Dim iJmlDataCeklis As Double
    Dim iJmlUser As Double
    Dim iMod As Double
        
    If optEven.Value = True Then
        'iJmlDataCeklis = Val(txtTotalDtaCek.Text)
        iJmlDataCeklis = tdbDataToDistribute.Value
        'iJmlDataCeklis = tdbDatarata.Value '--ELIN QUERY 21112019
        iJmlUser = Val(txtTotaltUser.Text)
        'iJmlUser = Val(txttotalspv.Text)
        ibg = 0
    
        iMod = 0
        If iJmlDataCeklis > 0 Then
            ibg = Fix(iJmlDataCeklis / iJmlUser)
            iMod = iJmlDataCeklis Mod iJmlUser
        End If
            
        If List_User.ListItems.Count > 0 Then
            For i = 1 To List_User.ListItems.Count
                iSisa = 0
                If iMod > 0 Then
                    iSisa = 1
                    iMod = iMod - iSisa
                End If
                List_User.ListItems(i).SubItems(3) = CStr(ibg) + iSisa
                'List_User.ListItems(i).SubItems(3) = tdbDatarata.Value / List_User.ListItems(i).SubItems(7)
                'List_User.ListItems(i).SubItems(3) = Int(List_User.ListItems(i).SubItems(3))
            Next i
        End If
        Call totalGetData
    End If
End Sub

Private Sub SSCommand1_Click()
    
End Sub

Private Sub tdbDataToDistribute_Change()
Dim a As String
On Error GoTo a
If MDIForm1.Txtlevel.Text = "Branch Manager" Then
    tdbDatarata.Value = 0
    a = Val(tdbDataToDistribute.Value) / Val(txttotalspv.Text)
    tdbDatarata.Value = Int(a)
End If
Exit Sub
a:
msg err.Description
End Sub

Private Sub tdbDataToDistribute_KeyPress(KeyAscii As Integer)
    If MDIForm1.Txtlevel.Text <> "Branch Manager" Then
        If KeyAscii = 13 Then
            If optEven.Value = True Then
                Call bagiRata
            End If
        End If
    End If
End Sub

Private Sub tdbDataToDistribute_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode >= 49 And KeyCode <= 57 Then
        If tdbDataToDistribute.Value > Val(txtTotalDtaCek.Text) Then
            msg ("number too big")
            KeyAscii = 0
            tdbDataToDistribute.Value = Val(txtTotalDtaCek.Text)
        End If
    End If
End Sub

'Private Sub Txt_Search_Change()
'    Call LoadCampaign(Cmb_BaseProduct, Txt_Search.Text)
'End Sub

Private Sub txtJmlAssign_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOk_Click
    End If
End Sub

Private Sub totalGetData()
    Dim sJ As Double
    
    sJ = 0
    If List_User.ListItems.Count > 0 Then
        For i = 1 To List_User.ListItems.Count
           sJ = sJ + Val(List_User.ListItems(i).SubItems(3))
        Next i
    End If
    
    txtTotalGetData.Text = CStr(sJ)
End Sub



Private Sub txtReasonException_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOKException_Click
    End If
End Sub
