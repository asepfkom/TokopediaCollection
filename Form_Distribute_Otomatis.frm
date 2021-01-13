VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Form_Distribute_Otomatis 
   BackColor       =   &H8000000E&
   Caption         =   "Distibute Otomatis"
   ClientHeight    =   9900
   ClientLeft      =   3120
   ClientTop       =   405
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   13935
   Begin VB.Frame Frame4 
      Height          =   9210
      Left            =   0
      TabIndex        =   3
      Top             =   675
      Width           =   13935
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
         Left            =   7290
         TabIndex        =   54
         Top             =   3660
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   390
         Left            =   6945
         TabIndex        =   53
         Top             =   6390
         Width           =   375
      End
      Begin VB.CommandButton cmdLoadUser 
         Caption         =   "Refresh User"
         Height          =   390
         Left            =   135
         TabIndex        =   50
         Top             =   8370
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdException 
         Caption         =   ">>"
         Height          =   390
         Left            =   6945
         TabIndex        =   30
         Top             =   5880
         Width           =   375
      End
      Begin VB.Frame frmUserAssign 
         Height          =   1305
         Left            =   1125
         TabIndex        =   29
         Top             =   4560
         Visible         =   0   'False
         Width           =   4815
         Begin Threed.SSCommand cmdCancel2 
            Height          =   345
            Left            =   4470
            TabIndex        =   49
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
            TabIndex        =   34
            Top             =   360
            Width           =   705
         End
         Begin VB.TextBox txtUserAssign 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1455
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   375
            Width           =   2325
         End
         Begin TDBNumber6Ctl.TDBNumber txtJmlAssign 
            Height          =   330
            Left            =   1455
            TabIndex        =   47
            Top             =   720
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   582
            Calculator      =   "Form_Distribute_Otomatis.frx":0000
            Caption         =   "Form_Distribute_Otomatis.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_Distribute_Otomatis.frx":008C
            Keys            =   "Form_Distribute_Otomatis.frx":00AA
            Spin            =   "Form_Distribute_Otomatis.frx":00F4
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
            TabIndex        =   33
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
            TabIndex        =   31
            Top             =   750
            Width           =   630
         End
      End
      Begin VB.CommandButton cmdProsesDistribusi 
         Caption         =   "Proses Distribusi"
         Height          =   390
         Left            =   5490
         TabIndex        =   28
         Top             =   8385
         Width           =   1440
      End
      Begin VB.Frame Frame6 
         Height          =   825
         Left            =   7080
         TabIndex        =   22
         Top             =   4035
         Width           =   6615
         Begin VB.TextBox txtTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0"
            Top             =   465
            Width           =   870
         End
         Begin VB.TextBox txtTotalDtaCek 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   5415
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0"
            Top             =   450
            Width           =   870
         End
         Begin TDBNumber6Ctl.TDBNumber tdbDataToDistribute 
            Height          =   330
            Left            =   1890
            TabIndex        =   52
            Top             =   435
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   582
            Calculator      =   "Form_Distribute_Otomatis.frx":011C
            Caption         =   "Form_Distribute_Otomatis.frx":013C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_Distribute_Otomatis.frx":01A8
            Keys            =   "Form_Distribute_Otomatis.frx":01C6
            Spin            =   "Form_Distribute_Otomatis.frx":0210
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
         Begin TDBNumber6Ctl.TDBNumber tdbDatarata 
            Height          =   330
            Left            =   345
            TabIndex        =   58
            Top             =   435
            Visible         =   0   'False
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   582
            Calculator      =   "Form_Distribute_Otomatis.frx":0238
            Caption         =   "Form_Distribute_Otomatis.frx":0258
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_Distribute_Otomatis.frx":02C4
            Keys            =   "Form_Distribute_Otomatis.frx":02E2
            Spin            =   "Form_Distribute_Otomatis.frx":032C
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
            TabIndex        =   59
            Top             =   135
            Visible         =   0   'False
            Width           =   1305
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
            Left            =   1560
            TabIndex        =   51
            Top             =   135
            Width           =   1665
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
            Left            =   4935
            TabIndex        =   41
            Top             =   135
            Width           =   1845
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
            Left            =   3300
            TabIndex        =   23
            Top             =   135
            Width           =   1635
         End
      End
      Begin VB.Frame Frame5 
         Height          =   405
         Left            =   4890
         TabIndex        =   19
         Top             =   3720
         Width           =   2055
         Begin VB.OptionButton optUneven 
            Appearance      =   0  'Flat
            Caption         =   "UnEven"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   21
            Top             =   150
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optEven 
            Appearance      =   0  'Flat
            Caption         =   "Even"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   75
            TabIndex        =   20
            Top             =   150
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
         TabIndex        =   13
         Top             =   4140
         Width           =   6840
         Begin VB.TextBox txttotalspv 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   56
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
            TabIndex        =   46
            Top             =   3765
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox txtTotalGetData 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3570
            Locked          =   -1  'True
            TabIndex        =   44
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
            TabIndex        =   24
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
            TabIndex        =   14
            Top             =   3765
            Visible         =   0   'False
            Width           =   1065
         End
         Begin MSComctlLib.ListView List_User 
            Height          =   3405
            Left            =   30
            TabIndex        =   15
            Top             =   285
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
            TabIndex        =   16
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
            TabIndex        =   57
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
            TabIndex        =   45
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
            TabIndex        =   25
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
            TabIndex        =   17
            Top             =   45
            Width           =   1245
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   0
            Top             =   15
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
         Height          =   4140
         Left            =   75
         TabIndex        =   4
         Top             =   165
         Width           =   13620
         Begin VB.TextBox txtJmlCampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   12795
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "0"
            Top             =   3495
            Width           =   780
         End
         Begin VB.CommandButton cmdLoadDataCampaign 
            Caption         =   "Load Data"
            Height          =   390
            Left            =   75
            TabIndex        =   18
            Top             =   3450
            Width           =   1440
         End
         Begin MSComctlLib.ListView List_Campaign 
            Height          =   3120
            Left            =   90
            TabIndex        =   6
            Top             =   315
            Width           =   13485
            _ExtentX        =   23786
            _ExtentY        =   5503
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
            Left            =   9645
            TabIndex        =   5
            Top             =   3525
            Visible         =   0   'False
            Width           =   1155
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   2760
            TabIndex        =   7
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
            TabIndex        =   27
            Top             =   3510
            Width           =   1425
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
            TabIndex        =   8
            Top             =   15
            Width           =   825
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
         TabIndex        =   9
         Top             =   5070
         Width           =   6735
         Begin VB.Frame frmConfirmExeption 
            Height          =   1305
            Left            =   555
            TabIndex        =   35
            Top             =   420
            Visible         =   0   'False
            Width           =   5490
            Begin Threed.SSCommand cmdCancel 
               Height          =   345
               Left            =   5190
               TabIndex        =   48
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
               TabIndex        =   38
               Top             =   570
               Width           =   2325
            End
            Begin VB.TextBox txtUserException 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   300
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   225
               Width           =   2325
            End
            Begin VB.CommandButton cmdOKException 
               Caption         =   "OK"
               Height          =   615
               Left            =   3885
               TabIndex        =   36
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
               TabIndex        =   40
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
               TabIndex        =   39
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
            TabIndex        =   10
            Top             =   3375
            Width           =   1455
         End
         Begin MSComctlLib.ListView List_Exception 
            Height          =   1605
            Left            =   255
            TabIndex        =   11
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
            TabIndex        =   12
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
         Left            =   1305
         TabIndex        =   55
         Top             =   3690
         Visible         =   0   'False
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
      Picture         =   "Form_Distribute_Otomatis.frx":0354
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
      Picture         =   "Form_Distribute_Otomatis.frx":0E5E
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
'Dim rs As ADODB.Recordset
'Dim sUser As String
'Dim sLevel As String
'Dim sExcepTrans As Boolean
'Dim isiAwalShowUser As String
'Private Sub CheckAll_Campaign_Click()
'    Call CheckAll(List_Campaign, CheckAll_Campaign)
'    Call itungJmlCek
'    Call bagiRata
'End Sub
'
'Private Sub CheckAll_Manager_Click()
'    Call CheckAll(List_User, CheckAll_Manager)
'    Call bagiRata
'End Sub
'Public Sub CheckAll(listview As listview, check As CheckBox)
'    Dim i As Integer
'    i = 0
'    If check.Value = 1 Then
'        For i = 1 To listview.ListItems.Count
'            listview.ListItems(i).Checked = True
'        Next i
'
'    ElseIf check.Value = 0 Then
'        For i = 1 To listview.ListItems.Count
'            listview.ListItems(i).Checked = False
'        Next i
'    End If
'
'End Sub
'
''Private Sub CheckAll_ProsesData_Click()
''    Call CheckAll(List_Proses_Data, CheckAll_ProsesData)
''End Sub
'
'Private Sub Cmb_BaseProduct_Change()
'    'Cmb_BaseProduct_Click
'End Sub
'
'Private Sub Cmb_BaseProduct_Click()
'    'Call loadDataCampaign
'    'Call ShowComboBox
'    Call loadDataUser
'    Call tampil_reason
'End Sub
'Public Sub ShowComboBox(Optional isiAwal As String)
'    sWhere = ""
'    Strsql = "select * from usertbl WHERE 1=1"
'
'    Select Case UCase(sLevel)
'        Case UCase("Supervisor")
'            sWhere = " and userid = '" & sUser & "' AND kdstatus ='1' AND f_cuti = 0"
'        Case UCase("Manager")
'            sWhere = " and userid in (select userid from usertbl where mgrcode = '" & sUser & "') AND kdstatus ='1' AND f_cuti = 0 AND usertype = '2'"
'        Case UCase("Branch Manager")
'            sWhere = " and userid in (select userid from usertbl where bm_code = '" & sUser & "') AND kdstatus ='1' AND f_cuti = 0 AND usertype = '5'"
'    End Select
'
'    Strsql = Strsql & sWhere
'
'    Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    Cmb_ShowUser.clear
'    Cmb_ShowUser.text = isiAwal
'    If isiAwal = "" Then
'        Cmb_ShowUser.text = sUserid
'        If UCase(sLevel) = "MANAGER" Then
'            Cmb_ShowUser.text = "[ALL]"
'        ElseIf UCase(sLevel) = "BRANCH MANAGER" Then
'            Cmb_ShowUser.text = "[agent]"
'        End If
'    End If
'
'    If UCase(sLevel) = "MANAGER" Then
'        Cmb_ShowUser.AddItem "[ALL]"
'        If MDIForm1.txt_f_cp.text = "0" Then
'            Cmb_ShowUser.AddItem "[ALL agent]"
'        End If
'    ElseIf UCase(sLevel) = "BRANCH MANAGER" Then
'        Cmb_ShowUser.AddItem "[ALL]"
'        Cmb_ShowUser.AddItem "[MANAGER]"
'        Cmb_ShowUser.AddItem "[SPV]"
'        Cmb_ShowUser.AddItem "[agent]"
'    End If
'    While Not rs.EOF
'        'Cmb_ShowUser.AddItem cnull(rs!userid)
'        Cmb_ShowUser.AddItem cnull(rs!Userid) & " - " & cnull(rs!agent)
'        rs.MoveNext
'    Wend
'
'    Set rs = Nothing
'End Sub
'Private Sub Cmb_ShowUser_Click()
'    Call loadDataUser
'    Call tampil_reason
'    isiAwalShowUser = Cmb_ShowUser.text
'End Sub
'
'Private Sub cmbCriteria_DropDown()
'    'Call getKriteria
'    cmbCriteria.clear
'    cmbCriteria.AddItem ""
'    cmbCriteria.AddItem "P1"
'    cmbCriteria.AddItem "P2"
'    cmbCriteria.AddItem "P3"
'    cmbCriteria.AddItem "P4"
'End Sub
'
'Private Sub getKriteria()
'    Dim rs As ADODB.Recordset
'    Dim Strsql, mwhere, strCamp As String
'
'    strCamp = getcampaign
'
'    If strCamp <> Empty Then
'        Strsql = " select distinct kriteria from mgm "
'        mwhere = " where 1=1"
'        mwhere = mwhere + " and recsource in (" + strCamp + ")  "
''    Else
''        msg "Pilih Campaign Terlebih Dahulu"
''        Exit Sub
'    End If
'
'    Strsql = Strsql + mwhere + " order by kriteria"
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    cmbCriteria.clear
'    While Not rs.EOF
'        cmbCriteria.AddItem cnull(rs!kriteria)
'        rs.MoveNext
'    Wend
'
'    Set rs = Nothing
'
'End Sub
'
''Private Sub getBaseProduct()
''    Dim rs As ADODB.Recordset
''    Dim STRSQL, MWHERE As String
''
''
''    STRSQL = " select keterangan from tblprogram"
''    MWHERE = " where 1=1"
''    Set rs = New ADODB.Recordset
''    rs.CursorLocation = adUseClient
''    rs.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
''
''    Cmb_BaseProduct.clear
''    While Not rs.EOF
''        Cmb_BaseProduct.AddItem cnull(rs!keterangan)
''        rs.MoveNext
''    Wend
''
''    Set rs = Nothing
''End Sub
'
'Private Sub Cmd_Load_Click()
'    Dim manager_boolean As Boolean
'    Dim campaign_boolean As Boolean
'    Dim i, jumlah_data, bagi_dua, ganjil, count_bagi, bagi_ke As Integer
'    List_Proses_Data.ColumnHeaders.clear
'    List_Proses_Data.ColumnHeaders.ADD 1, , "Campaign Code", 2500
'    Cmd_Load.Enabled = False
'    manager_boolean = False
'    count_bagi = 0
'    K = 1
'    For i = 1 To List_Manager.ListItems.Count
'        If List_Manager.ListItems(i).Checked = True Then
'            List_Proses_Data.ColumnHeaders.ADD K + 1, , List_Manager.ListItems(i).text
'            count_bagi = count_bagi + 1
'            K = K + 1
'            manager_boolean = True
'        End If
'    Next i
'    If manager_boolean = False Or K > 3 Then
'        'Call msg("Manager Tidak Boleh Kosong Atau Tidak Boleh Dari 2")
'        Exit Sub
'    End If
'    campaign_boolean = False
'    bagi_ke = 1
'    List_Proses_Data.ListItems.clear
'
'    For i = 1 To List_Campaign.ListItems.Count
'        If List_Campaign.ListItems(i).Checked = True Then
'            jumlah_data = List_Campaign.ListItems(i).SubItems(1)
'            bagi_dua = Int(jumlah_data / count_bagi)
'            ganjil = jumlah_data - (bagi_dua * count_bagi)
'
'
'            Set listv = List_Proses_Data.ListItems.ADD(1, , List_Campaign.ListItems(i).text)
'            For J = 1 To count_bagi
'                If ganjil = 1 Then
'                    If bagi_ke = J Then
'                        listv.SubItems(J) = bagi_dua + 1
'                        If bagi_ke = 2 Then
'                            bagi_ke = 1
'                        Else
'                            bagi_ke = 2
'                        End If
'                        ganjil = 0
'                    Else
'                        listv.SubItems(J) = bagi_dua
'                        ganjil = 1
'                    End If
'                Else
'                    listv.SubItems(J) = bagi_dua
'                End If
'            Next J
'            campaign_boolean = True
'        End If
'    Next i
'    If campaign_boolean = False Then
'        'Call msg("Campaign Code Tidak Boleh Kosong")
'        Exit Sub
'    End If
'    Cmd_Proses.Enabled = True
'    Cmd_Load.Enabled = True
'End Sub
'
'Private Sub cmd_proses_Click()
'    Dim proses_data As Boolean
'    Dim RECSOURCE, user, Nama, jumlah_data, sQueryid As String
'    Cmd_Proses.Enabled = False
'    Cmd_Load.Enabled = False
'    proses_data = False
'    For i = 1 To List_Proses_Data.ListItems.Count
'        If List_Proses_Data.ListItems(1).Checked = True Then
'            proses_data = True
'            sWhere = ""
'            sWhere = cTable.CheckWhere(sWhere, "(agent is null or agent='' or agent='" & MDIForm1.Text7.text & "')")
'            sWhere = cTable.CheckWhere(sWhere, "statuscall='New Data'")
'
'            RECSOURCE = List_Proses_Data.ListItems(1).text
'            sWhere = cTable.CheckWhere(sWhere, "recsource = '" & RECSOURCE & "'")
'
'            For J = 1 To List_Proses_Data.ColumnHeaders.Count - 1
'                user = List_Proses_Data.ColumnHeaders(J + 1).text
'                Nama = get_name_user(user)
'                jumlah_data = List_Proses_Data.ListItems(1).SubItems(J)
'                sQueryid = "SELECT id FROM mgm " & sWhere & " ORDER BY id LIMIT " & jumlah_data
'                DoEvents
'                'Call CreateInsert_Waterfall_Hst(sQueryid, user, nama, MDIForm1.Text2.Text)
'
'                sQueryUpdate = " UPDATE mgm " & _
'                               " SET agent='" + user + "', nmagent= '" + Nama + "', tgl_distribusi =date(now()) " & _
'                               " WHERE id in( " & sQueryid & " ) and statuscall <> 'Agree'  "
'                DoEvents
'                M_OBJCONN.execute sQueryUpdate
'
'                sQueryInsert = " INSERT INTO tbllogdistribusi(userid,nama,recsource,jmldata,sendby) values " & _
'                               " ('" & user & "'," & _
'                               " '" & Nama & "'," & _
'                               " '" & RECSOURCE & "'," & _
'                               CStr(Val(jumlah_data)) & "," & _
'                               " '" + MDIForm1.Text7.text + "')"
'                DoEvents
'                M_OBJCONN.execute sQueryInsert
'            Next J
'            List_Proses_Data.ListItems.Remove (1)
'        End If
'    Next i
'    Cmd_Load.Enabled = True
'    If proses_data = False Then
'        msg ("Silahkan Pilih List Proses Data")
'        Cmd_Proses.Enabled = True
'        Exit Sub
'    End If
'    msg ("Distribusi Otomatis Sukses")
'    CheckAll_Campaign.Value = vbUnchecked
'    CheckAll_ProsesData.Value = vbUnchecked
'    Cmb_BaseProduct_Click
'    List_Proses_Data.ListItems.clear
'
'End Sub
'Function get_name_user(sUserid)
'    Call UnConnectRs(rs)
'    sQuerySelect = "SELECT agent FROM usertbl WHERE userid='" & sUserid & "'"
'    rs.Open sQuerySelect
'    If rs.EOF Then
'        get_name_user = ""
'        Exit Function
'    End If
'    get_name_user = cnull(rs!agent)
'End Function
'
'Private Sub CreateInsert_Waterfall_Hst(sQueryid As String, ByVal agent As String, ByVal nmAgent As String, Level As String)
'    Dim sBulan      As String
'    Dim sYear       As String
'    Dim sTableagent  As String
'
'    sBulan = Format(FungsiWaktuServer, "mmmm")
'    sYear = Format(FungsiWaktuServer, "yyyy")
'    sTableagent = "tbl_mgm_hst_" & sBulan & "_" & sYear
'    On Error GoTo InsertTable
'    'Call CreateTableMgmHst
'InsertTable:
'    Level = UCase(Level)
'    If Level = "agent" Then
'        sQuerySelect = " select id as id_cust,'New Data'::text,'New Data'::text,'" & agent & "'::text,'" & nmAgent & "'::text,spvcode,ketgroupspv,now(),recsource,campaign_agent from mgm a,usertbl b where a.agent=b.userid " & _
'                       " and id in ( " & sQueryid & " ) and statuscall <> 'Agree'  "
'    Else
'        sQuerySelect = " select id as id_cust,'New Data'::text,'New Data'::text,'" & agent & "'::text,'" & nmAgent & "'::text,'" & agent & "'::text,'" & nmAgent & "'::text,now(),recsource,campaign_agent from mgm where  " & _
'                       " id in ( " & sQueryid & " ) and statuscall <> 'Agree'  "
'    End If
'
'    sQueryInsert = " INSERT INTO " & sTableagent & "(" & _
'                   " id_cust,statuscall,reasoncall,agent,nmagent,kdspv,nmspv,tglcall,recsource,campaign_agent " & _
'                   " ) " & sQuerySelect
'    M_OBJCONN.execute sQueryInsert
'End Sub
'
'Private Sub CmdCancel_Click()
'    frmConfirmExeption.Visible = False
'End Sub
'
'Private Sub cmdCancel2_Click()
'    frmUserAssign.Visible = False
'End Sub
'
'Private Sub cmdException_Click()
'    Dim list As listItem
'
'    With List_Exception
'        .ColumnHeaders.clear
'        .ColumnHeaders.ADD , , "No.", 5 * TXT
'        .ColumnHeaders.ADD , , "Userid", 10 * TXT
'        .ColumnHeaders.ADD , , "Nama", 0 * TXT
'        .ColumnHeaders.ADD , , "Reason", 30 * TXT
'    End With
'
'    If List_User.ListItems.Count = 0 Then
'        Exit Sub
'    End If
'    For i = 1 To List_Exception.ListItems.Count
'        If List_Exception.ListItems(i).SubItems(1) = List_User.SelectedItem.SubItems(1) Then
'            MsgBox "Data ini Telah Dipilih", vbInformation, "Informasi"
'            Exit Sub
'        End If
'    Next i
'    txtUserException.text = List_User.SelectedItem.SubItems(1)
'    sExcepTrans = True
'    frmConfirmExeption.Visible = True
'    txtReasonException.SetFocus
'End Sub
'
'Private Sub cmdOKException_Click()
'
'    If sExcepTrans = True Then
'        Set lst = List_Exception.ListItems.ADD(, , List_User.SelectedItem.text)
'        lst.SubItems(1) = List_User.SelectedItem.SubItems(1)
'        lst.SubItems(2) = List_User.SelectedItem.SubItems(2)
'        lst.SubItems(3) = txtReasonException.text
'        List_User.ListItems.Remove (List_User.SelectedItem.Index)
'        Call itungJmlUser
'        If txtTotaltUser <> "0" Then
'            Call bagiRata
'        End If
'        Call insertLogException
'        Call resetException
'        Call ShowComboBox(isiAwalShowUser)
'        Call loadDataUser
'        Call tampil_reason
'    Else
'        List_Exception.SelectedItem.SubItems(3) = txtReasonException.text
'        Call resetException
'    End If
'End Sub
'
'Private Sub cmdProsesDistribusi_Click()
'    If txtTotaltUser.text = "0" Then
'        MsgBox "Tidak ada User", vbCritical + vbOKOnly, "TINS"
'        Exit Sub
'    End If
'    If Val(txtTotalGetData.text) > Val(txtTotalDtaCek.text) Then
'        If MDIForm1.Text2.text = "Branch Manager" Then
'            If MsgBox("Jumlah data(" + txtTotalDtaCek.text + ") kurang dari jumlah data(" + txtTotalGetData.text + ") yang akan di distribusi,yakin akan di proses?", vbQuestion + vbYesNo, "TINS") = vbNo Then
'                Exit Sub
'            End If
'        Else
'            MsgBox "Jumlah data(" + txtTotalDtaCek.text + ") kurang dari jumlah data(" + txtTotalGetData.text + ") yang akan di distribusi", vbCritical + vbOKOnly, "TINS"
'            Exit Sub
'        End If
'    End If
'
'    If Val(txtTotalGetData.text) < Val(Replace(tdbDataToDistribute.text, ",", "")) Then
'        If MDIForm1.Text2.text = "Branch Manager" Then
'            If MsgBox("Jumlah data(" + txtTotalGetData.text + ") kurang dari jumlah data(" + tdbDataToDistribute.text + ") yang akan di distribusi,yakin akan di proses?", vbQuestion + vbYesNo, "TINS") = vbNo Then
'                Exit Sub
'            End If
'        Else
'            MsgBox "Jumlah data(" + txtTotalGetData.text + ") kurang dari jumlah data(" + tdbDataToDistribute.text + ") yang akan di distribusi", vbCritical + vbOKOnly, "TINS"
'            Exit Sub
'        End If
'    End If
'
'    If Val(txtTotalGetData.text) > Val(Replace(tdbDataToDistribute.text, ",", "")) Then
'        If MDIForm1.Text2.text = "Branch Manager" Then
'            If MsgBox("Jumlah data(" + txtTotalGetData.text + ") lebih dari jumlah data(" + tdbDataToDistribute.text + ") yang akan di distribusi,yakin akan di proses?", vbQuestion + vbYesNo, "TINS") = vbNo Then
'                Exit Sub
'            End If
'        Else
'            MsgBox "Jumlah data(" + txtTotalGetData.text + ") lebih dari jumlah data(" + tdbDataToDistribute.text + ") yang akan di distribusi", vbCritical + vbOKOnly, "TINS"
'            Exit Sub
'        End If
'    End If
'
'    cmdProsesDistribusi.Enabled = False
'    Call Prosesdistribusi
'    cmdProsesDistribusi.Enabled = True
'End Sub
'
'Private Sub insertLogException()
'    Dim cmdsql, sUserException, sExceptionReason As String
'
'        cmdsql = " insert into tbl_log_distribute_exception (userid,nama,exception_reason,userid_execute) "
'        cmdsql = cmdsql + " values ('" + txtUserException.text + "','" + List_Exception.SelectedItem.SubItems(2) + "','" + txtReasonException.text + "','" + sUser + "');" + vbCrLf
'        M_OBJCONN.execute cmdsql
'
'        cmdsql = "update usertbl set f_cuti='1', exception_reason = '" + txtReasonException.text + "' where userid='" + txtUserException.text + "'"
'        M_OBJCONN.execute cmdsql
'End Sub
'Private Sub insertLogDistribusi()
'    Dim cmdsql, sUserDistrib, sExceptionReason As String
'    On Error GoTo ERRORA
'
'    cmdsql = Empty
'    If List_User.ListItems.Count > 0 Then
'        For i = 1 To List_User.ListItems.Count
'            sUserDistrib = List_User.ListItems(i).SubItems(1)
'            sNamaDistrib = List_User.ListItems(i).SubItems(2)
'            sCampaignCde = getcampaign("1")
'            sJmlDataDist = List_User.ListItems(i).SubItems(3)
'
'            cmdsql = cmdsql + " insert into tbllogdistribusi (userid,nama,recsource,jmldata,sendby) "
'            cmdsql = cmdsql + " values ('" + sUserDistrib + "','" + sNamaDistrib + "','" + sCampaignCde + "','" + sJmlDataDist + "','" + sUser + "');" + vbCrLf
'        Next i
'        M_OBJCONN.execute cmdsql
'    End If
'
'Exit Sub
'ERRORA:
'    MsgBox "Error Loging!:" + err.Description, vbInformation + vbOKOnly, "TINS"
'End Sub
'
'Private Sub itungJmlUser()
'    txtTotaltUser.text = CStr(List_User.ListItems.Count)
'End Sub
'
'Private Function getcampaign(Optional sCode As String) As String
'    Dim strCampaign As String
'    Dim sC As String
'
'    sC = IIf(sCode = "1", "", "'")
'    For i = 1 To List_Campaign.ListItems.Count
'        If List_Campaign.ListItems(i).Checked = True Then
'            If strCampaign = Empty Then
'                strCampaign = strCampaign + sC + List_Campaign.ListItems(i).SubItems(1) + sC
'            Else
'                strCampaign = strCampaign + "," + sC + List_Campaign.ListItems(i).SubItems(1) + sC + ""
'            End If
'        End If
'    Next i
'
'    getcampaign = strCampaign
'End Function
'
'Private Sub Prosesdistribusi()
'On Error GoTo KE
'    Dim rs As ADODB.Recordset
'    Dim IJ, adk As Integer
'    Dim Strsql, strCampaign, sUserPil, sLmit, SELECT_LOG, CRT_LOG As String
'
'    If List_Campaign.ListItems.Count = 0 Then
'        MsgBox "Pilih Campaign Terlebih dahulu", vbOKOnly + vbCritical, "TINS"
'        Exit Sub
'    End If
'
'    If Val(txtTotalGetData.text) = 0 Then
'        MsgBox "Isi jumlah data yang akan didistribusi", vbCritical + vbOKOnly, "TINS"
'        Exit Sub
'    End If
'
'    strCampaign = getcampaign
'    If strCampaign = Empty Then
'        MsgBox "Pilih Campaign Terlebih Dahulu", vbCritical + vbOKOnly, "TINS"
'        Exit Sub
'    End If
'
'    adk = 0
'    For IJ = 1 To List_Campaign.ListItems.Count
'        If List_Campaign.ListItems(IJ).Checked = True Then
'            adk = adk + 1
'        End If
'    Next IJ
'
'    If adk > 1 Then
'        MsgBox ("untuk sementara pembagian data hanya bisa per satu campaign")
'        Exit Sub
'    End If
'
'    If MsgBox("Anda yakin ingin melakukan distribusi?", vbQuestion + vbYesNo, "TINS") = vbNo Then
'        Exit Sub
'    End If
'
'    For i = 1 To List_User.ListItems.Count
'        sUserPil = List_User.ListItems(i).SubItems(1)
'        sLmit = CStr(Val(List_User.ListItems(i).SubItems(3)))
'
'        If Val(sLmit) > 0 Then
'            SELECT_LOG = " select id from mgm where 1=1 "
'            SELECT_LOG = SELECT_LOG + vbCrLf + " and recsource in (" + strCampaign + ") "
'            'SELECT_LOG = SELECT_LOG + vbCrLf + " and agent = '" + sUser + "' "
'            SELECT_LOG = SELECT_LOG + vbCrLf + " and coalesce(agent,'') = '' "
'            'SELECT_LOG = SELECT_LOG + vbCrLf + " and coalesce(f_block,0) = '0' "
'            SELECT_LOG = SELECT_LOG + vbCrLf + " and recsource not in (select distinct kodeds from datasourcetbl where status = '0')"
'            '--block campaign FIRAS
''            If MDIForm1.Text2.text = "Supervisor" Then
''                SELECT_LOG = SELECT_LOG & " and recsource not in (select distinct recsource from tblblockcampaign_tmp where agent ='" & MDIForm1.Text7.text & "' and status = '5' and exec_user in (select mgrcode from usertbl where userid = '" & MDIForm1.Text7.text & "'))"
''            End If
'            '--
''            If cmbCriteria.text <> Empty Then
''                SELECT_LOG = SELECT_LOG + vbCrLf + " and kriteria = '" + cmbCriteria.text + "' "
''            End If
''
''            If Cmb_BaseProduct.text <> Empty Then
''                SELECT_LOG = SELECT_LOG + vbCrLf + " and base_product = '" + Cmb_BaseProduct.text + "' "
''            End If
'
'            SELECT_LOG = SELECT_LOG + vbCrLf + " order by row_number() over(partition by recsource) "
'            SELECT_LOG = SELECT_LOG + vbCrLf + " limit " + sLmit
'
'            cmdsql = "update mgm set agent = '" + sUserPil + "', nmagent = '" & List_User.ListItems(i).SubItems(2) & "',tgl_distribusi= now() where id in (" + SELECT_LOG + " ); "
'            '===================================================================
'            CRT_LOG = "insert into tbl_log_distribute_auto (to_userid,from_userid,id_mgm) "
'            CRT_LOG = CRT_LOG + vbCrLf + " select '" + sUserPil + "','" + sUser + "',id from (" + SELECT_LOG + ") as a;"
'
'            M_OBJCONN.execute cmdsql + CRT_LOG
'        End If
'    Next i
'
'
'    Call insertLogDistribusi
'    MsgBox "Proses Distribusi Done", vbInformation + vbOKOnly, "TINS"
'
'    loadDataCampaign
'    loadDataUser
'Exit Sub
'KE:
'MsgBox err.Description, vbInformation + vbOKOnly, "TINS"
'
'End Sub
'
'Private Sub Command1_Click()
'    If List_Exception.ListItems.Count <> 0 Then
'        Set lList = List_User.ListItems.ADD(, , List_Exception.SelectedItem.text)
'        lList.SubItems(1) = List_Exception.SelectedItem.SubItems(1)
'        lList.SubItems(2) = List_Exception.SelectedItem.SubItems(2)
'        List_Exception.ListItems.Remove List_Exception.SelectedItem.Index
'
'        cmdsql = "update usertbl set f_cuti='0', exception_reason = '' where userid='" + lList.SubItems(1) + "'        "
'        M_OBJCONN.execute cmdsql
'
'        cmdsql = "update tbl_log_distribute_exception set flag='1' where userid='" + lList.SubItems(1) + "'        "
'        M_OBJCONN.execute cmdsql
'        Call ShowComboBox(isiAwalShowUser)
'        Call loadDataUser
'        Call tampil_reason
'    End If
'End Sub
'
'Private Sub Form_Load()
'    Me.Width = 14175
'    Me.Height = 10935
'    sExcepTrans = False
'    sUser = MDIForm1.Text7.text
'    sLevel = MDIForm1.Text2.text
'    'Call getBaseProduct
'    If MDIForm1.txt_f_cp.text = "0" Then
'        Cmb_BaseProduct.clear
'        Cmb_BaseProduct.text = "TOPUP PL"
'        Cmb_BaseProduct.Enabled = False
'    ElseIf MDIForm1.txt_f_cp.text = "1" Then
'        Cmb_BaseProduct.clear
'        Cmb_BaseProduct.text = "MORTGAGE"
'        Cmb_BaseProduct.AddItem "MORTGAGE"
'        Cmb_BaseProduct.AddItem "KPM"
'        Cmb_BaseProduct.Enabled = True
'    End If
'
'
'    If MDIForm1.Text2.text = "Branch Manager" Then
'        Cmb_BaseProduct.text = "TOKOPEDIA"
'        Cmb_BaseProduct.AddItem "TOKOPEDIA"
'        Cmb_BaseProduct.Enabled = True
'        Cmb_ShowUser.AddItem "[MANAGER]"
'        Cmb_ShowUser.AddItem "[SUPERVISOR]"
'        cmbflag.Visible = True
'        Label1(2).Visible = True
'        Label18.Visible = True
'        tdbDatarata.Visible = True
'        Label17.Visible = True
'        txttotalspv.Visible = True
'        Call getFlag
'    End If
'    Call ShowComboBox
'    Call loadDataUser
'    Call tampil_reason
'End Sub
'Private Sub getFlag()
'    Dim rs As ADODB.Recordset
'    Dim Strsql, mwhere, strCamp As String
'
'
'    If Cmb_BaseProduct.text = "MORTGAGE" Then
'        Strsql = " select distinct f_param from usertbl where f_cp='1' and f_kpm='0'"
'    ElseIf Cmb_BaseProduct.text = "KPM" Or Cmb_BaseProduct.text = "TOKOPEDIA" Then
'        Strsql = " select distinct f_param from usertbl where f_cp='1' and f_kpm='1'"
'    ElseIf Cmb_BaseProduct.text = "TOPUP PL" Then
'        Strsql = " select distinct f_param from usertbl where f_cp='0'"
'    End If
'
'    Strsql = Strsql + mwhere + " order by f_param"
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    cmbflag.clear
'    While Not rs.EOF
'        cmbflag.AddItem cnull(rs!f_param)
'        rs.MoveNext
'    Wend
'
'    Set rs = Nothing
'
'End Sub
'Public Sub UnConnectRs(ByRef rsS As ADODB.Recordset)
'    On Error Resume Next
'    If rsS.state = 1 Then rsS.Close
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    Call UnConnectRs(rs)
'End Sub
'Private Sub cmdLoadDataCampaign_Click()
'    cmdLoadDataCampaign.Enabled = False
'    Call loadDataCampaign
'    Call loadDataUser
'    cmdLoadDataCampaign.Enabled = True
'End Sub
'Private Sub loadDataCampaign()
'    Dim rs As ADODB.Recordset
'    Dim list As listItem
'    Dim Strsql, sUseragent, sField As String
'    Dim MHWERE As String
'    Dim sJ As Double
'
'
'    Strsql = "select"
'    Strsql = Strsql + vbCrLf + " recsource"
'    Strsql = Strsql + vbCrLf + " ,sum(case when coalesce(agent,'') = '' then 1 else 0 end) as remain"
'    Strsql = Strsql + vbCrLf + " ,count(m.id) as total_data"
'    'STRSQL = STRSQL + vbCrLf + " ,sum(case when agent='" + sUser + "' then 1 else 0 end) as recycle"
'    Strsql = Strsql + vbCrLf + " from mgm m"
'    Strsql = Strsql + vbCrLf + " left join datasourcetbl d on (d.kodeds=m.recsource)"
'
'    mwhere = vbCrLf + " Where 1 = 1"
'    'mwhere = mwhere + vbCrLf + " and date(tglexpire)>date(now())"
''    MWHERE = MWHERE + vbCrLf + " and (date(startdate)<=date(now()) or startdate is null)"
''    MWHERE = MWHERE + vbCrLf + " and coalesce(m.f_block,0) = '0'"
'    'MWHERE = MWHERE + vbCrLf + " and d.kdstatus = '1'"
'
'    If MDIForm1.Text2.text = "Manager" Then
'        mwhere = mwhere + vbCrLf + " and (d.officer_id = '" & MDIForm1.Text7.text & "' )"
'    End If
'
'    sField = ""
'    If sLevel = "Supervisor" Then
'        sField = "spvcode"
'    Else
'        sField = "userid"
'    End If
'
'    'MWHERE = MWHERE + vbCrLf + " and (agent ='" + sUser + "' or agent in (select userid from usertbl where " + sField + "='" + sUser + "'))"
'
'
''    If Txt_Search.text <> Empty Then
''        MWHERE = MWHERE + vbCrLf + " and datasourcetbl_recsource  like '%" + Txt_Search.text + "%'"
''    End If
'
'    Strsql = Strsql + mwhere + vbCrLf + " group by recsource order by recsource"
'
'    'Call UnConnectRs(rs)
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    txtJmlCampaign.text = CStr(Val(rs.RecordCount))
'
'    If rs.RecordCount = 0 Then
'        MsgBox "Data Not Found", vbInformation + vbOKOnly, "TINS"
'        Exit Sub
'    End If
'
'    With List_Campaign
'        .ListItems.clear
'        .ColumnHeaders.clear
'        .ColumnHeaders.ADD , , "No.", 5 * TXT
'        .ColumnHeaders.ADD , , "Campaign Code", 55 * TXT
'        .ColumnHeaders.ADD , , "Remain", 10 * TXT
'        .ColumnHeaders.ADD , , "Total Data", 10 * TXT
'        '.ColumnHeaders.ADD , , "Recycle", 10 * TXT
'
'        .ListItems.clear
'
'        sJ = 0
'        While Not rs.EOF
'            Set list = .ListItems.ADD(, , rs.Bookmark)
'            list.SubItems(1) = cnull(rs!RECSOURCE)
'            list.SubItems(2) = cnull(rs!Remain)
'            list.SubItems(3) = cnull(rs!total_data)
'            'list.SubItems(4) = cnull(rs!recycle)
'            sJ = sJ + Val(cnull(rs!total_data))
'            rs.MoveNext
'        Wend
'        txtTotal.text = CStr(Val(sJ))
'        'tdbDataToDistribute.Value = Val(sJ)
'    End With
'    Call itungJmlCek
'    Set rs = Nothing
'End Sub
'Private Sub cmdLoadUser_Click()
'    cmdLoadUser.Enabled = False
'    Call loadDataUser
'    Call tampil_reason
'    cmdLoadUser.Enabled = True
'End Sub
'
'Private Sub loadDataUser()
'    Dim rs As ADODB.Recordset
'    Dim list As listItem
'    Dim Strsql, strCampaign As String
'    Dim MHWERE As String
'    Dim getCode As String
'    Dim getagent As String
'
'    '=================================================
'    getCode = ""
'    getagent = ""
'    txttotalspv.text = ""
'    intvrl = InStr(1, Cmb_ShowUser.text, " - ", vbTextCompare)
'    If intvrl <> 0 Then
'       ArrayString = Split(Cmb_ShowUser.text, " - ", 2, vbTextCompare)
'       getCode = ArrayString(0)
'       getagent = ArrayString(1)
'    End If
'    '=================================================
'
'    strCampaign = getcampaign
'    If MDIForm1.Text2.text = "Branch Manager" Then
'        Strsql = " select "
'        Strsql = Strsql + vbCrLf + "  userid As USERID"
'        Strsql = Strsql + vbCrLf + "  ,agent as nama"
'        Strsql = Strsql + vbCrLf + "  ,u.spvcode as SPV"
'        Strsql = Strsql + vbCrLf + "  ,total_agent"
'        Strsql = Strsql + vbCrLf + "  ,total_spv"
'    Else
'        Strsql = " select "
'        Strsql = Strsql + vbCrLf + "  userid As USERID"
'        Strsql = Strsql + vbCrLf + "  ,agent as nama"
'    End If
'
'    If strCampaign <> Empty And chkShowAssign.Value = 1 Then
'        Strsql = Strsql + vbCrLf + "  ,coalesce(a.jml_distributed,'0') as jml_distributed"
'    End If
'
'    If MDIForm1.Text2.text = "Branch Manager" Then
'        Strsql = Strsql + vbCrLf + " from usertbl u"
'        Strsql = Strsql + vbCrLf + " LEFT JOIN (select count(distinct userid) as total_agent,spvcode from usertbl where usertype = '1'"
'        Strsql = Strsql + vbCrLf + " and kdstatus = '1' group by spvcode) t on (u.spvcode=t.spvcode)"
'        Strsql = Strsql + vbCrLf + " LEFT JOIN (select count(distinct spvcode) as total_spv,mgrcode from usertbl where usertype = '2'"
'        Strsql = Strsql + vbCrLf + " and kdstatus = '1' group by mgrcode) v on (u.mgrcode=v.mgrcode)"
'        'strsql = strsql + vbCrLf + " inner join (select distinct userid from tblloguseradm where date(tgl)=date(now()) and operation='Success Login') w on (u.userid=w.userid)"
'    Else
'        Strsql = Strsql + vbCrLf + " from usertbl u"
'    End If
'
'    If strCampaign <> Empty And chkShowAssign.Value = 1 Then
'        Strsql = Strsql + vbCrLf + " left join ("
'        Strsql = Strsql + vbCrLf + "    select"
'
'        If sLevel = "Supervisor" Then
'            Strsql = Strsql + vbCrLf + " agent as user_group"
'        End If
'
'        Strsql = Strsql + vbCrLf + "    ,count(id) as jml_distributed"
'        Strsql = Strsql + vbCrLf + "    from mgm m"
'        Strsql = Strsql + vbCrLf + "    left join usertbl u on (u.userid=m.agent)"
'        Strsql = Strsql + vbCrLf + "    Where 1 = 1"
'        Strsql = Strsql + vbCrLf + "    and recsource in (" + strCampaign + ") "
'        Strsql = Strsql + vbCrLf + "group by user_group"
'        Strsql = Strsql + vbCrLf + ")as a on (a.user_group=u.userid)"
'    End If
'
'    If sLevel = "Supervisor" Then
'        mwhere = " Where f_cuti='0'"
'        'MWHERE = MWHERE + vbCrLf + " and spvcode ='" + sUser + "'"
'        mwhere = mwhere + vbCrLf + " and usertype = '1'"
'        'MWHERE = MWHERE + vbCrLf + " and kdstatus = '1'"
'    End If
'    '=================================================
'
'    Strsql = Strsql + mwhere + " order by 1"
'
'    'Call UnConnectRs(rs)
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    txtTotaltUser.text = CStr(Val(rs.RecordCount))
'    If rs.RecordCount <> 0 Then
'        If MDIForm1.Text2.text = "Branch Manager" Then
'            txttotalspv.text = cnull(rs!total_spv)
'        End If
'
'
'    With List_User
'        .ListItems.clear
'        .ColumnHeaders.clear
'        .ColumnHeaders.ADD , , "No.", 5 * TXT
'        .ColumnHeaders.ADD , , "userid", 10 * TXT
'        .ColumnHeaders.ADD , , "agent", 20 * TXT
'        .ColumnHeaders.ADD , , "Data to Distribute", 10 * TXT
'        .ColumnHeaders.ADD , , "Data Assign", 10 * TXT
'        .ColumnHeaders.ADD , , "Flag", 10 * TXT
'        If MDIForm1.Text2.text = "Branch Manager" Then
'            .ColumnHeaders.ADD , , "SPV", 10 * TXT
'            .ColumnHeaders.ADD , , "Jml agent", 10 * TXT
'            .ColumnHeaders.ADD , , "Jml SPV", 10 * TXT
'        End If
'        .ListItems.clear
''        While Not rs.EOF
''            Set list = .ListItems.Add(, , rs.Bookmark)
''            list.SubItems(1) = cnull(rs!UserId)
''            list.SubItems(2) = cnull(rs!nama)
''            list.SubItems(3) = "0"
''            If strCampaign <> Empty And chkShowAssign.Value = 1 Then
''                list.SubItems(4) = cnull(rs!jml_distributed)
''            End If
''            rs.MoveNext
''        Wend
'
'        While Not rs.EOF
'            Set list = .ListItems.ADD(, , rs.Bookmark)
'            list.SubItems(1) = cnull(rs!Userid)
'            If Mid(list.SubItems(1), 1, 5) = "TUPSP" Or Mid(list.SubItems(1), 1, 5) = "OMSPV" Then
'                list.ListSubItems(1).ForeColor = vbRed
'            Else
'                list.ListSubItems(1).ForeColor = vbBlack
'            End If
'            list.SubItems(2) = cnull(rs!Nama)
'            list.SubItems(3) = "0"
'            If strCampaign <> Empty And chkShowAssign.Value = 1 Then
'                list.SubItems(4) = cnull(rs!jml_distributed)
'            End If
'            'list.SubItems(5) = cnull(rs!f_param)
'            If MDIForm1.Text2.text = "Branch Manager" Then
'                list.SubItems(6) = cnull(rs!spv)
'                list.SubItems(7) = cnull(rs!total_agent)
'                list.SubItems(8) = cnull(rs!total_spv)
'            End If
'            rs.MoveNext
'        Wend
'        'txtJmlDatatoUser.Text = "0"
'    End With
'    End If
'
'    Set rs = Nothing
'
'End Sub
'Private Sub loadDataUser_spv()
'    Dim rs As ADODB.Recordset
'    Dim list As listItem
'    Dim Strsql, strCampaign As String
'    Dim MHWERE As String
'    Dim getCode As String
'    Dim getagent As String
'
'    '=================================================
'    If MDIForm1.Text2.text = "Branch Manager" Then
'        Strsql = " select distinct spv as spv1,jmlperspv from tbl_tmp_distribution_by_spv"
'    End If
'    '=================================================
'
'    'Call UnConnectRs(rs)
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'    If rs.RecordCount <> 0 Then
'    With ListView1
'        .ListItems.clear
'        .ColumnHeaders.clear
'        .ColumnHeaders.ADD , , "No.", 5 * TXT
'        .ColumnHeaders.ADD , , "SPV", 10 * TXT
'        .ColumnHeaders.ADD , , "Jumlah", 20 * TXT
'        .ListItems.clear
'        While Not rs.EOF
'            Set list = .ListItems.ADD(, , rs.Bookmark)
'            list.SubItems(1) = cnull(rs!spv1)
'            list.SubItems(2) = cnull(rs!jmlperspv)
'            rs.MoveNext
'        Wend
'    End With
'    End If
'
'    Set rs = Nothing
'
'End Sub
'Private Sub tampil_reason()
'Dim rs As ADODB.Recordset
'Dim Strsql As String
'
'    Strsql = "select * from usertbl where mgrcode='" + MDIForm1.Text1.text + "' and usertype = '1' and f_cuti='1'"
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'    no = 0
'     With List_Exception
'        .ColumnHeaders.clear
'        .ColumnHeaders.ADD , , "No.", 5 * TXT
'        .ColumnHeaders.ADD , , "Userid", 10 * TXT
'        .ColumnHeaders.ADD , , "Nama", 0 * TXT
'        .ColumnHeaders.ADD , , "Reason", 30 * TXT
'    End With
'    List_Exception.ListItems.clear
'    While Not rs.EOF
'    no = no + 1
'        Set list = List_Exception.ListItems.ADD(, , no)
'            list.SubItems(1) = cnull(rs!Userid)
'            list.SubItems(2) = cnull(rs!agent)
'            list.SubItems(3) = cnull(rs!exception_reason)
'        rs.MoveNext
'    Wend
'    Set rs = Nothing
'End Sub
'Private Sub List_Campaign_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    If ColumnHeader.Index = 1 Then 'ANGKA
'        If List_Campaign.SortOrder = 0 Then
'            Call SortColumn(List_Campaign, ColumnHeader.Index, sortDescending, sortNumeric)
'        Else
'            Call SortColumn(List_Campaign, ColumnHeader.Index, sortAscending, sortNumeric)
'        End If
'    Else
'        If List_Campaign.SortOrder = 0 Then 'HURUF
'            Call SortColumn(List_Campaign, ColumnHeader.Index, sortDescending, sortAlpha)
'        Else
'            Call SortColumn(List_Campaign, ColumnHeader.Index, sortAscending, sortAlpha)
'        End If
'    End If
'End Sub
'
'Private Sub List_Campaign_ItemCheck(ByVal Item As MSComctlLib.listItem)
'    Call itungJmlCek
'    Call bagiRata
'End Sub
'
'Private Sub itungJmlCek()
''On Error GoTo a
'    Dim sJ As Double
'    Dim i As Integer
'
'    sJ = 0
'    For i = 1 To List_Campaign.ListItems.Count
'        If List_Campaign.ListItems(i).Checked = True Then
'            sJ = sJ + Val(List_Campaign.ListItems(i).SubItems(2))
'        End If
'    Next i
'    txtTotalDtaCek.text = CStr(Val(sJ))
'    tdbDataToDistribute.Value = Val(sJ)
'    If MDIForm1.Text2.text = "Branch Manager" Then
'        If tdbDataToDistribute.Value <> 0 Then
'            tdbDatarata.Value = tdbDataToDistribute.Value / txttotalspv.text
'        Else
'            tdbDatarata.Value = 0
'        End If
'    End If
'Exit Sub
''a:
''msg err.Description
'End Sub
'
'Private Sub List_Exception_DblClick()
'    frmConfirmExeption.Visible = True
'    txtUserException.text = List_Exception.SelectedItem.SubItems(1)
'    txtReasonException.text = List_Exception.SelectedItem.SubItems(2)
'
'End Sub
'
'Private Sub List_User_Click()
'        Call resetException
'End Sub
'Private Sub resetException()
'    txtUserException.text = Empty
'    txtReasonException.text = Empty
'    frmConfirmExeption.Visible = False
'    sExcepTrans = False
'End Sub
'
'Private Sub List_User_DblClick()
'Dim rs As ADODB.Recordset
'Dim Strsql, mwhere, strCampaign As String
'
''strCampaign = getcampaign
''If strCampaign <> Empty Then
''    STRSQL = " select * from datasourcetbl "
''    MWHERE = " where 1=1"
''    MWHERE = MWHERE + " and datasourcetbl_recsource in (" + strCampaign + ")  "
''
''    STRSQL = STRSQL + MWHERE
''
''    Set rs = New ADODB.Recordset
''    rs.CursorLocation = adUseClient
''    rs.open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
''
''    If UCase(cnull(rs!distribusitype)) = "UNEVEN" Then
'        If optUneven.Value = True Then
'            txtUserAssign.text = List_User.SelectedItem.SubItems(1)
'            txtJmlAssign.Value = Val(List_User.SelectedItem.SubItems(3))
'            frmUserAssign.Visible = True
'            txtJmlAssign.SetFocus
'        End If
''    End If
''
''    Set rs = Nothing
''End If
'End Sub
'Private Sub CmdOK_Click()
'    List_User.SelectedItem.SubItems(3) = CStr(Val(txtJmlAssign.text))
'    txtJmlAssign.text = Empty
'    frmUserAssign.Visible = False
'    List_User.SetFocus
'    Call totalGetData
'End Sub
'
'Private Sub List_User_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        List_User_DblClick
'    End If
'End Sub
'
'Private Sub optEven_Click()
'    If txtTotaltUser.text = "0" Then
'        MsgBox "Tidak ada User", vbCritical + vbOKOnly, "TINS"
'        Exit Sub
'    Else
'        If MDIForm1.Text2.text = "Branch Manager" Then
'            Call Bagi_rata_per_spv
'            Call totalGetData
'            Call loadDataUser_spv
'        Else
'            Call bagiRata
'            Call totalGetData
'        End If
'    End If
'End Sub
'
'Sub Bagi_rata_per_spv()
'On Error GoTo a
'    M_OBJCONN.execute "Delete from tbl_tmp_distribution_by_spv"
'    For J = 1 To List_User.ListItems.Count
'
'      xx = "  insert into  tbl_tmp_distribution_by_spv(agent,spv, jumlah_agent,jmlperspv) values ('" & List_User.ListItems(J).SubItems(1) & "','" & List_User.ListItems(J).SubItems(6) & "','" & List_User.ListItems(J).SubItems(7) & "','" & Val(tdbDatarata) & "')"
'        M_OBJCONN.execute xx
'    Next J
'
'    DoEvents
'    M_OBJCONN.execute "Select * from sp_bagirata_perspv ()"
'    For i = 1 To List_User.ListItems.Count
'        Strsql = "select * from tbl_tmp_distribution_by_spv where agent='" + List_User.ListItems(i).SubItems(1) + "'"
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'        If rs.RecordCount <> 0 Then
'            List_User.ListItems(i).SubItems(3) = cnull(rs!jmldata)
'        End If
'    Next i
'    Exit Sub
'a:
'    msg err.Description
'End Sub
'Private Sub optUneven_Click()
'    Call totalGetData
'End Sub
'
'Private Sub bagiRata()
'    Dim ibg As Integer
'    Dim iJmlDataCeklis As Double
'    Dim iJmlUser As Double
'    Dim iMod As Double
'
'    If optEven.Value = True Then
'        'iJmlDataCeklis = Val(txtTotalDtaCek.Text)
'        iJmlDataCeklis = tdbDataToDistribute.Value
'        'iJmlDataCeklis = tdbDatarata.Value '--ELIN QUERY 21112019
'        iJmlUser = Val(txtTotaltUser.text)
'        'iJmlUser = Val(txttotalspv.Text)
'        ibg = 0
'
'        iMod = 0
'        If iJmlDataCeklis > 0 Then
'            ibg = Fix(iJmlDataCeklis / iJmlUser)
'            iMod = iJmlDataCeklis Mod iJmlUser
'        End If
'
'        If List_User.ListItems.Count > 0 Then
'            For i = 1 To List_User.ListItems.Count
'                iSisa = 0
'                If iMod > 0 Then
'                    iSisa = 1
'                    iMod = iMod - iSisa
'                End If
'                List_User.ListItems(i).SubItems(3) = CStr(ibg) + iSisa
'                'List_User.ListItems(i).SubItems(3) = tdbDatarata.Value / List_User.ListItems(i).SubItems(7)
'                'List_User.ListItems(i).SubItems(3) = Int(List_User.ListItems(i).SubItems(3))
'            Next i
'        End If
'        Call totalGetData
'    End If
'End Sub
'
'Private Sub SSCommand1_Click()
'
'End Sub
'
'Private Sub tdbDataToDistribute_Change()
'Dim a As String
''On Error GoTo a
'If MDIForm1.Text2.text = "Branch Manager" Then
'    tdbDatarata.Value = 0
'    a = Val(tdbDataToDistribute.Value) / Val(txttotalspv.text)
'    tdbDatarata.Value = Int(a)
'End If
'Exit Sub
''a:
''msg err.Description
'End Sub
'
'Private Sub tdbDataToDistribute_KeyPress(KeyAscii As Integer)
'    If MDIForm1.Text2.text <> "Branch Manager" Then
'        If KeyAscii = 13 Then
'            If optEven.Value = True Then
'                Call bagiRata
'            End If
'        End If
'    End If
'End Sub
'
'Private Sub tdbDataToDistribute_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode >= 49 And KeyCode <= 57 Then
'        If tdbDataToDistribute.Value > Val(txtTotalDtaCek.text) Then
'            msg ("number too big")
'            KeyAscii = 0
'            tdbDataToDistribute.Value = Val(txtTotalDtaCek.text)
'        End If
'    End If
'End Sub
'
''Private Sub Txt_Search_Change()
''    Call LoadCampaign(Cmb_BaseProduct, Txt_Search.Text)
''End Sub
'
'Private Sub txtJmlAssign_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        CmdOK_Click
'    End If
'End Sub
'
'Private Sub totalGetData()
'    Dim sJ As Double
'
'    sJ = 0
'    If List_User.ListItems.Count > 0 Then
'        For i = 1 To List_User.ListItems.Count
'           sJ = sJ + Val(List_User.ListItems(i).SubItems(3))
'        Next i
'    End If
'
'    txtTotalGetData.text = CStr(sJ)
'End Sub
'
'
'
'Private Sub txtReasonException_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        cmdOKException_Click
'    End If
'End Sub
'
'
