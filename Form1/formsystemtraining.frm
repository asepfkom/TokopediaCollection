VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form formsystemtraining 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000015&
   BorderStyle     =   0  'None
   Caption         =   "Training Form"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20460
   LinkTopic       =   "Form5"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   11520
   ScaleWidth      =   20460
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   11535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   20346
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Display Setting"
      TabPicture(0)   =   "formsystemtraining.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "formsystemtraining.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control"
      TabPicture(2)   =   "formsystemtraining.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         Height          =   11055
         Left            =   120
         TabIndex        =   120
         Top             =   360
         Width           =   20295
         Begin VB.Frame Frame6 
            BackColor       =   &H0080FF80&
            Caption         =   "REPORT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6495
            Left            =   7800
            TabIndex        =   140
            Top             =   4320
            Width           =   12135
            Begin VB.CommandButton Command9 
               Caption         =   "Export"
               Height          =   375
               Left            =   5400
               TabIndex        =   147
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Search"
               Height          =   375
               Left            =   4080
               TabIndex        =   141
               Top             =   240
               Width           =   1215
            End
            Begin MSComctlLib.ListView ListView3 
               Height          =   5505
               Left            =   120
               TabIndex        =   142
               Top             =   720
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   9710
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Left            =   840
               TabIndex        =   143
               Top             =   300
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   556
               Calendar        =   "formsystemtraining.frx":0054
               Caption         =   "formsystemtraining.frx":016C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "formsystemtraining.frx":01D8
               Keys            =   "formsystemtraining.frx":01F6
               Spin            =   "formsystemtraining.frx":0254
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
               Format          =   "dd-mm-yyyy"
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
               Text            =   "__-__-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37468
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate2 
               Height          =   315
               Left            =   2565
               TabIndex        =   144
               Top             =   300
               Width           =   1305
               _Version        =   65536
               _ExtentX        =   2302
               _ExtentY        =   556
               Calendar        =   "formsystemtraining.frx":027C
               Caption         =   "formsystemtraining.frx":0394
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "formsystemtraining.frx":0400
               Keys            =   "formsystemtraining.frx":041E
               Spin            =   "formsystemtraining.frx":047C
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
               Format          =   "dd-mm-yyyy"
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
               Text            =   "__-__-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37468
               CenturyMode     =   0
            End
            Begin MSComDlg.CommonDialog CD_save 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "To"
               Height          =   255
               Index           =   3
               Left            =   2280
               TabIndex        =   146
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   145
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   3855
            Left            =   7800
            TabIndex        =   134
            Top             =   240
            Width           =   12135
            Begin VB.CommandButton Command15 
               Caption         =   "Refresh"
               Height          =   375
               Left            =   10680
               TabIndex        =   138
               Top             =   0
               Width           =   1215
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   3345
               Left            =   120
               TabIndex        =   137
               Top             =   360
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   5900
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   10575
            Left            =   360
            TabIndex        =   122
            Top             =   240
            Width           =   6975
            Begin VB.CommandButton Command8 
               Caption         =   "Set"
               Height          =   375
               Left            =   5520
               TabIndex        =   136
               Top             =   9960
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Check All"
               Height          =   255
               Left            =   120
               TabIndex        =   127
               Top             =   10200
               Width           =   5175
            End
            Begin VB.ComboBox Combo1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               TabIndex        =   123
               Top             =   240
               Width           =   3735
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   8145
               Left            =   120
               TabIndex        =   125
               Top             =   2160
               Width           =   5205
               _ExtentX        =   9181
               _ExtentY        =   14367
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin TDBDate6Ctl.TDBDate tgl1 
               Height          =   315
               Left            =   120
               TabIndex        =   129
               Top             =   1200
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   556
               Calendar        =   "formsystemtraining.frx":04A4
               Caption         =   "formsystemtraining.frx":05BC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "formsystemtraining.frx":0628
               Keys            =   "formsystemtraining.frx":0646
               Spin            =   "formsystemtraining.frx":06A4
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
               Format          =   "dd-mm-yyyy"
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
               Text            =   "__-__-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37468
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate tgl2 
               Height          =   315
               Left            =   3045
               TabIndex        =   130
               Top             =   1200
               Width           =   1305
               _Version        =   65536
               _ExtentX        =   2302
               _ExtentY        =   556
               Calendar        =   "formsystemtraining.frx":06CC
               Caption         =   "formsystemtraining.frx":07E4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "formsystemtraining.frx":0850
               Keys            =   "formsystemtraining.frx":086E
               Spin            =   "formsystemtraining.frx":08CC
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
               Format          =   "dd-mm-yyyy"
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
               Text            =   "__-__-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37468
               CenturyMode     =   0
            End
            Begin TDBTime6Ctl.TDBTime jam1 
               Height          =   375
               Left            =   1560
               TabIndex        =   131
               Top             =   1200
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               Caption         =   "formsystemtraining.frx":08F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "formsystemtraining.frx":0960
               Spin            =   "formsystemtraining.frx":09B0
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__:__"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   0.507210648148148
            End
            Begin TDBTime6Ctl.TDBTime jam2 
               Height          =   375
               Left            =   4440
               TabIndex        =   132
               Top             =   1200
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               Caption         =   "formsystemtraining.frx":09D8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "formsystemtraining.frx":0A44
               Spin            =   "formsystemtraining.frx":0A94
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__:__"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   0.507210648148148
            End
            Begin VB.Label Label7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5640
               TabIndex        =   139
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label5 
               Height          =   255
               Left            =   5400
               TabIndex        =   135
               Top             =   240
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "to"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   2535
               TabIndex        =   133
               Top             =   1320
               Width           =   225
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Set Time"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   128
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Select Participant"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   126
               Top             =   1920
               Width           =   5175
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "List Training Data"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   124
               Top             =   270
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   10935
         Left            =   -74880
         TabIndex        =   117
         Top             =   360
         Width           =   20175
         Begin VB.CommandButton Command7 
            Caption         =   "Prev"
            Height          =   375
            Left            =   17160
            TabIndex        =   121
            Top             =   10200
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Next"
            Height          =   375
            Left            =   18600
            TabIndex        =   119
            Top             =   10200
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            Height          =   9615
            Left            =   120
            ScaleHeight     =   9555
            ScaleWidth      =   19875
            TabIndex        =   118
            Top             =   240
            Width           =   19935
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         Height          =   11535
         Left            =   -75000
         TabIndex        =   1
         Top             =   360
         Width           =   20415
         Begin VB.CommandButton Command4 
            Caption         =   "Clear Pic"
            Height          =   375
            Left            =   15960
            TabIndex        =   149
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Save 
            Caption         =   "Save"
            Height          =   375
            Left            =   13320
            TabIndex        =   61
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Exit"
            Height          =   375
            Left            =   18600
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox Text1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            TabIndex        =   59
            Top             =   210
            Width           =   2895
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Load"
            Height          =   255
            Left            =   4440
            TabIndex        =   58
            Top             =   240
            Width           =   615
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   8
            Left            =   17520
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   57
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   0
            Left            =   120
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   56
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   1
            Left            =   2280
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   55
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   2
            Left            =   4440
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   2115
            TabIndex        =   54
            Top             =   840
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   3
            Left            =   6720
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   53
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   4
            Left            =   8880
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   52
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   5
            Left            =   11040
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   51
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   6
            Left            =   13200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   50
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   7
            Left            =   15360
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   49
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   9
            Left            =   120
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   48
            Top             =   2520
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   10
            Left            =   2280
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   47
            Top             =   2520
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   11
            Left            =   4440
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   2115
            TabIndex        =   46
            Top             =   2520
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   12
            Left            =   6720
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   45
            Top             =   2520
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   13
            Left            =   8880
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   44
            Top             =   2520
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   14
            Left            =   11040
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   43
            Top             =   2520
            Width           =   2055
         End
         Begin VB.CommandButton Command5 
            Caption         =   "View"
            Height          =   375
            Left            =   14640
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   15
            Left            =   13200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   41
            Top             =   2520
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   16
            Left            =   15360
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   40
            Top             =   2520
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   17
            Left            =   17520
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   39
            Top             =   2520
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   18
            Left            =   120
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   38
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   19
            Left            =   2280
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   37
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   20
            Left            =   4440
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   2115
            TabIndex        =   36
            Top             =   4200
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   21
            Left            =   6720
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   35
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   22
            Left            =   8880
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   34
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   23
            Left            =   11040
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   33
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   24
            Left            =   13200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   32
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   25
            Left            =   15360
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   31
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   26
            Left            =   17520
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   30
            Top             =   4200
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   27
            Left            =   120
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   29
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   28
            Left            =   2280
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   28
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   29
            Left            =   4440
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   2115
            TabIndex        =   27
            Top             =   5880
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   30
            Left            =   6720
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   26
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   31
            Left            =   8880
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   25
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   32
            Left            =   11040
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   24
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   33
            Left            =   13200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   23
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   34
            Left            =   15360
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   22
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   35
            Left            =   17520
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   21
            Top             =   5880
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   36
            Left            =   120
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   20
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   37
            Left            =   2280
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   19
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   38
            Left            =   4440
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   2115
            TabIndex        =   18
            Top             =   7560
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   39
            Left            =   6720
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   17
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   40
            Left            =   8880
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   16
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   41
            Left            =   11040
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   15
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   42
            Left            =   13200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   14
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   43
            Left            =   15360
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   13
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   44
            Left            =   17520
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   12
            Top             =   7560
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   45
            Left            =   120
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   11
            Top             =   9240
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   46
            Left            =   2280
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   10
            Top             =   9240
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   47
            Left            =   4440
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   2115
            TabIndex        =   9
            Top             =   9240
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   48
            Left            =   6720
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   8
            Top             =   9240
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   49
            Left            =   8880
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   7
            Top             =   9240
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   50
            Left            =   11040
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   6
            Top             =   9240
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   51
            Left            =   13200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   5
            Top             =   9240
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   52
            Left            =   15360
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   4
            Top             =   9240
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Index           =   53
            Left            =   17520
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1395
            ScaleWidth      =   1995
            TabIndex        =   3
            Top             =   9240
            Width           =   2055
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Delete"
            Height          =   375
            Left            =   17280
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   120
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label8 
            Caption         =   "Label8"
            Height          =   375
            Left            =   5880
            TabIndex        =   148
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   116
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   115
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   114
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   7320
            TabIndex        =   113
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   4
            Left            =   9480
            TabIndex        =   112
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   5
            Left            =   11760
            TabIndex        =   111
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   6
            Left            =   13920
            TabIndex        =   110
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   7
            Left            =   16080
            TabIndex        =   109
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   18240
            TabIndex        =   108
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   107
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   2760
            TabIndex        =   106
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   5160
            TabIndex        =   105
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   12
            Left            =   7320
            TabIndex        =   104
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   9480
            TabIndex        =   103
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   11760
            TabIndex        =   102
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "List"
            Height          =   255
            Left            =   720
            TabIndex        =   101
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   15
            Left            =   13920
            TabIndex        =   100
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   16
            Left            =   16080
            TabIndex        =   99
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   17
            Left            =   18240
            TabIndex        =   98
            Top             =   4080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   18
            Left            =   720
            TabIndex        =   97
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   19
            Left            =   2760
            TabIndex        =   96
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   20
            Left            =   5160
            TabIndex        =   95
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   7200
            TabIndex        =   94
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   9480
            TabIndex        =   93
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   23
            Left            =   11640
            TabIndex        =   92
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   24
            Left            =   13920
            TabIndex        =   91
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   25
            Left            =   15960
            TabIndex        =   90
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   26
            Left            =   18120
            TabIndex        =   89
            Top             =   5760
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   27
            Left            =   720
            TabIndex        =   88
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   28
            Left            =   2760
            TabIndex        =   87
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   29
            Left            =   5160
            TabIndex        =   86
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   30
            Left            =   7200
            TabIndex        =   85
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   31
            Left            =   9480
            TabIndex        =   84
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   32
            Left            =   11520
            TabIndex        =   83
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   33
            Left            =   13800
            TabIndex        =   82
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   34
            Left            =   15840
            TabIndex        =   81
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   35
            Left            =   18120
            TabIndex        =   80
            Top             =   7440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   36
            Left            =   720
            TabIndex        =   79
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   37
            Left            =   2760
            TabIndex        =   78
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   38
            Left            =   5160
            TabIndex        =   77
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   39
            Left            =   7200
            TabIndex        =   76
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   40
            Left            =   9480
            TabIndex        =   75
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   41
            Left            =   11520
            TabIndex        =   74
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   42
            Left            =   13800
            TabIndex        =   73
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   43
            Left            =   15840
            TabIndex        =   72
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   44
            Left            =   18120
            TabIndex        =   71
            Top             =   9120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   45
            Left            =   720
            TabIndex        =   70
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   46
            Left            =   2760
            TabIndex        =   69
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   47
            Left            =   5160
            TabIndex        =   68
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   48
            Left            =   7200
            TabIndex        =   67
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   49
            Left            =   9480
            TabIndex        =   66
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   50
            Left            =   11520
            TabIndex        =   65
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   51
            Left            =   13800
            TabIndex        =   64
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   52
            Left            =   15840
            TabIndex        =   63
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   53
            Left            =   18240
            TabIndex        =   62
            Top             =   10800
            Visible         =   0   'False
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "formsystemtraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ii As Integer

Private Sub Check1_Click()
    Dim K As Integer
       
    For K = 1 To ListView1.ListItems.Count
        LvPTP.ListItems(K).Checked = True
    Next K
End Sub


Private Sub Combo1_Click()
    Strsql = "SELECT * From tblsystemtraining where nama_file = '" & Combo1.text & "';"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    Label5.Caption = rs!ID
    
End Sub

Private Sub Combo1_DropDown()
    Strsql = "SELECT * From tblsystemtraining;"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    Combo1.clear
    
    If rs.RecordCount <> 0 Then
        For i = 1 To rs.RecordCount
            Combo1.AddItem rs!nama_file
            rs.MoveNext
        Next i
    End If

End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command15_Click()
    Call isilv2
    Call JAM
End Sub

Private Sub Command18_Click()
    Call isilv3
End Sub

Private Sub Command2_Click()
    Dim aaa As String
    Dim bbb As String
    Dim i As Integer
    Dim S As Integer
    
    'clear pic
    For i = 0 To 53
        Set Picture1(i).Picture = Nothing
    Next i
    
    'getpath
    'aaa = "C:\" & Text1.text & "\"
    aaa = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & Text1.text & "\"
    
    'load
    For i = 0 To 53
        If i <= 53 Then
            bbb = aaa & i & ".jpg"
            
            If CheckPath(bbb) = True Then
                Picture1(i).ScaleMode = 3
                Picture1(i).AutoRedraw = True
                Picture1(i).Picture = LoadPicture(bbb)
                Picture1(i).PaintPicture Picture1(i).Picture, _
                0, 0, Picture1(i).ScaleWidth, Picture1(i).ScaleHeight, _
                0, 0, Picture1(i).Picture.Width / 26.46, _
                Picture1(i).Picture.Height / 26.46
                Picture1(i).Picture = Picture1(i).Image
            End If
        End If
    Next i
    
End Sub

Private Sub Command3_Click()
    Dim aaa As String
    Dim bbb As String

    ii = ii + 1

    'getpath
    'aaa = "C:\" & Text1.text & "\"
    aaa = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & Text1.text & "\"
        
    'load
    If ii <= 53 Then
        bbb = aaa & ii & ".jpg"
        
        If CheckPath(bbb) = True Then
            Picture2.ScaleMode = 3
            Picture2.AutoRedraw = True
            Picture2.Picture = LoadPicture(bbb)
            Picture2.PaintPicture Picture2.Picture, _
            0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, _
            0, 0, Picture2.Picture.Width / 26.46, _
            Picture2.Picture.Height / 26.46
            Picture2.Picture = Picture2.Image
'        Else
'            MsgBox "Training sudah selesai"
'            Unload Me
        End If
    End If
End Sub

Private Sub Command4_Click()
    'Frame2.Visible = False
    For i = 0 To 53
        Set Picture1(i).Picture = Nothing
    Next i
End Sub

Private Sub headeragentnparticipant()
    ListView1.ColumnHeaders.ADD 1, , "AGENT", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "NAMA AGENT", 70 * TXT
    ListView1.ColumnHeaders.ADD 3, , "TEAM", 10 * TXT
    
    ListView2.ColumnHeaders.ADD 1, , "TRAINING", 20 * TXT
    ListView2.ColumnHeaders.ADD 2, , "JAM AWAL", 20 * TXT
    ListView2.ColumnHeaders.ADD 3, , "JAM AKHIR", 20 * TXT
    ListView2.ColumnHeaders.ADD 4, , "AGENT", 20 * TXT
    ListView2.ColumnHeaders.ADD 5, , "NAMA", 20 * TXT
    ListView2.ColumnHeaders.ADD 6, , "STATUS", 20 * TXT
    'ListView2.ColumnHeaders.ADD 7, , "SIGN", 20 * TXT
    
    ListView3.ColumnHeaders.ADD 1, , "TRAINING", 20 * TXT
    ListView3.ColumnHeaders.ADD 2, , "TANGGAL", 20 * TXT
    ListView3.ColumnHeaders.ADD 3, , "AGENT", 20 * TXT
    ListView3.ColumnHeaders.ADD 4, , "NAMA", 20 * TXT
    ListView3.ColumnHeaders.ADD 5, , "STATUS", 20 * TXT
    ListView3.ColumnHeaders.ADD 6, , "SIGN", 20 * TXT
    
End Sub

Private Sub isilv2()
    Dim listItem As listItem
 
    qsel = "select * from ("
    qsel = qsel & " select nama_file as training, jam_awal, jam_akhir, agent, nama ,f_done, b.ids  from tblsystemtraining_schedule a inner join"
    qsel = qsel & " (select a.*, b.agent as nama from tblsystemtraining_partisipan a left join usertbl b on a.agent = b.userid) b on a.ids = b.ids inner join"
    qsel = qsel & " tblsystemtraining c on a.idp = c.id"
    qsel = qsel & " ) a where jam_awal < now() and jam_akhir > now() order by jam_awal, agent"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qsel, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListView2.ListItems.clear
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Set listItem = ListView2.ListItems.ADD(, , cnull(rs!training))
                listItem.SubItems(1) = Format(rs!jam_awal, "yyyy-mm-dd hh:nn:ss")
                listItem.SubItems(2) = Format(rs!jam_akhir, "yyyy-mm-dd hh:nn:ss")
                listItem.SubItems(3) = cnull(rs!agent)
                listItem.SubItems(4) = cnull(rs!Nama)
                If cnull(rs!f_done) <> "" Then
                    listItem.SubItems(5) = "COMPLETE"
                Else
                    listItem.SubItems(5) = ""
                End If
            rs.MoveNext
        Next i
    End If

End Sub

Private Sub isilv3()
    Dim listItem As listItem
 
    qsel = "select * from ("
    qsel = qsel & " select nama_file as training, jam_awal, jam_akhir, agent, nama ,f_done, b.ids  from tblsystemtraining_schedule a inner join"
    qsel = qsel & " (select a.*, b.agent as nama from tblsystemtraining_partisipan a left join usertbl b on a.agent = b.userid) b on a.ids = b.ids inner join"
    qsel = qsel & " tblsystemtraining c on a.idp = c.id"
    qsel = qsel & " ) a "
    
    If IsNull(TDBDate1.Value) = False And IsNull(TDBDate2.Value) = False Then
        j1 = Format(TDBDate1.Value, "yyyy-mm-dd")
        j2 = Format(TDBDate2.Value, "yyyy-mm-dd")
        
        qsel = qsel & " where jam_awal between '" & j1 & "' and '" & j2 & "' "
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qsel, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListView3.ListItems.clear
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Set listItem = ListView3.ListItems.ADD(, , cnull(rs!training))
                listItem.SubItems(1) = Format(rs!jam_awal, "yyyy-mm-dd")
                listItem.SubItems(2) = cnull(rs!agent)
                listItem.SubItems(3) = cnull(rs!Nama)
                If cnull(rs!f_done) <> "" Then
                    listItem.SubItems(4) = "COMPLETE"
                Else
                    listItem.SubItems(4) = ""
                End If
                'listItem.SubItems(5) = cnull(rs!Nama)
            rs.MoveNext
        Next i
    End If

End Sub


Private Sub isilv()
    Dim listItem As listItem
    qsel = "select userid,agent,team from usertbl where aktif = 0 and usertype in (1,6) and (spvcode <> 'RESERVED' and agent <> 'DECEASE') order by 3,1;"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qsel, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListView1.ListItems.clear
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Set listItem = ListView1.ListItems.ADD(, , cnull(rs!Userid))
                listItem.SubItems(1) = cnull(rs!agent)
                listItem.SubItems(2) = cnull(rs!TEAM)
            rs.MoveNext
        Next i
    End If
End Sub


Private Sub Command5_Click()
    Dim aaa As String
    Dim bbb As String
    Dim i As Integer
    Dim S As Integer
    
'    Frame2.Top = 1320
'    Frame2.Left = 2880
    Frame2.Visible = True
    
    SSTab1.Tab = 1
     
    'getpath
    'aaa = "C:\" & Text1.text & "\"
    aaa = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & Text1.text & "\"
    
    ii = 0
    
    'load
    If ii <= 53 Then
        bbb = aaa & ii & ".jpg"
        
        If CheckPath(bbb) = True Then
            Picture2.ScaleMode = 3
            Picture2.AutoRedraw = True
            Picture2.Picture = LoadPicture(bbb)
            Picture2.PaintPicture Picture2.Picture, _
            0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, _
            0, 0, Picture2.Picture.Width / 26.46, _
            Picture2.Picture.Height / 26.46
            Picture2.Picture = Picture2.Image
        End If
    End If
    
End Sub

Private Sub Command6_Click()
    Dim zzz As String
    
    If Label8.Caption <> "Label8" Then
        'clear pic
        For i = 0 To 53
            Set Picture1(i).Picture = Nothing
        Next i
        
        qdel = "delete from tblsystemtraining where id = " & Label8.Caption
        M_OBJCONN.execute qdel
        
        'delete directory
        'zzz = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & Text1.text & "\"
        'RmDir (zzz)
        
        MsgBox "Berhasil di delete"
        Label8.Caption = "Label8"
    End If
    
End Sub

Private Sub Command7_Click()
    Dim aaa As String
    Dim bbb As String

    'getpath
    'aaa = "C:\" & Text1.text & "\"
    aaa = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & Text1.text & "\"
        
    'load
    If ii >= 1 And ii <= 53 Then
        ii = ii - 1
        bbb = aaa & ii & ".jpg"
        
        If CheckPath(bbb) = True Then
            Picture2.ScaleMode = 3
            Picture2.AutoRedraw = True
            Picture2.Picture = LoadPicture(bbb)
            Picture2.PaintPicture Picture2.Picture, _
            0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, _
            0, 0, Picture2.Picture.Width / 26.46, _
            Picture2.Picture.Height / 26.46
            Picture2.Picture = Picture2.Image
        End If
    End If

End Sub

Private Sub Command8_Click()
    Dim jaw As String
    Dim jak As String
    Dim idsch As Integer
    Dim a As Integer
    Dim qins As String
    a = 0
    
    If tgl1.ValueIsNull = True And tgl2.ValueIsNull = True Then
        MsgBox "Harap Pilih Tanggal"
        Exit Sub
    End If
    
    If jam1.ValueIsNull = True And jam2.ValueIsNull = True Then
        MsgBox "Harap Pilih Jam"
        Exit Sub
    End If
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            a = a + 1
        End If
    Next i
    
    If a = 0 Then
        MsgBox "Harap Pilih Participant"
        Exit Sub
    End If
    
    If Label5.Caption = "" Then
        MsgBox "Harap Pilih Data Training"
        Exit Sub
    End If
    
    'tblsystemtraining.id = tblsystemtraining_schedule.idp = tblsystemtraining_partisipan.ids
    
    jaw = Format(tgl1, "yyyy-mm-dd") & " " & jam1.Value
    jak = Format(tgl2, "yyyy-mm-dd") & " " & jam2.Value
    
    qins = "insert into tblsystemtraining_schedule (jam_awal,jam_akhir,idp) values ('" & jaw & "','" & jak & "'," & Label5.Caption & ");" & vbCrLf
    M_OBJCONN.execute qins
    
    Strsql = "SELECT * From tblsystemtraining_schedule order by ids desc limit 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    idsch = rs!ids
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            qins = "insert into tblsystemtraining_partisipan (agent,ids) values ('" & ListView1.ListItems(i).text & "', " & idsch & ");"
            qins = qins & "insert into tblsystemtraining_partisipan_log (agent,ids) values ('" & ListView1.ListItems(i).text & "', " & idsch & ");"
            M_OBJCONN.execute qins
        End If
    Next i
    
    Strsql = "SELECT * From tblsystemtraining_partisipan"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount >= 1000 Then
        qdel = "delete from tblsystemtraining_partisipan where idme not in (select idme from tblsystemtraining_partisipan order by idme desc limit 1000)"
        M_OBJCONN.execute qdel
    End If
    
    MsgBox "Berhasil di Set"
End Sub

Private Sub Command9_Click()
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

Private Sub save_Click()
    Dim zzz As String
    Dim yyy As String
    Dim S As Integer
    Dim ifilenumber As Integer
    Static iErrCtr As Integer
    
    iErrCtr = iErrCtr + 1
    ifilenumber = FreeFile
    
    If Text1.text = "" Then
        MsgBox "Harap isi Nama Slide"
        Exit Sub
    End If
    
    'zzz = "C:\" & Text1.text & "\"
    zzz = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & Text1.text & "\"
    
    If CheckPath(zzz) = False Then
        'simpandb
        qins = "insert into tblsystemtraining (nama_file,addby) values ('" & Text1.text & "', '" & MDIForm1.Text1.text & "');"
        M_OBJCONN.execute qins

        yyy = zzz
        MkDir (yyy)
        Open yyy & "README.txt" For Append As #ifilenumber
            Write #ifilenumber, "AAA"
        Close #ifilenumber
        
        'hitunggambar
        S = 0
        For i = 0 To 53
            If Picture1(i).Picture <> 0 Then
                S = S + 1
            End If
        Next i
        
        
        'savegambar
        For i = 0 To 53
            If Picture1(i).Picture <> 0 Then
                yyy = zzz
                yyy = yyy & i & ".jpg"
                If CheckPath(yyy) = False Then
                    'SavePicture Picture1(i).Image, yyy
                    FileCopy Label1(i).Caption, yyy
                End If
            End If
        Next i
            
    Else
        'hitunggambar
        S = 0
        For i = 0 To 53
            If Picture1(i).Picture <> 0 Then
                S = S + 1
            End If
        Next i
        
        'savegambar
        For i = 0 To 53
            If Picture1(i).Picture <> 0 Then
                yyy = zzz
                yyy = yyy & i & ".jpg"
                If CheckPath(yyy) = False Then
                    'SavePicture Picture1(i).Image, yyy
                    FileCopy Label1(i).Caption, yyy
                End If
            End If
        Next i
    End If
    
    MsgBox "Data Berhasil Disimpan"
End Sub

Private Sub JAM()
    Strsql = "select to_char(now(),'hh24:mi') as jamskrg"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    Label7.Caption = rs!jamskrg
    
    
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    Call cektbl
    
    If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Then
        Frame2.Visible = False
    End If
    
    Call JAM
    
    Call headeragentnparticipant
    Call isilv
    Call isilv2
    
'    Picture1(0).Picture = LoadPicture("D:\CC.jpg")
'    For i = 0 To 14
'        Picture1(i).ScaleMode = 3
'        Picture1(i).AutoRedraw = True
'        'Picture1(i).AutoSize = True
'        'Picture1(i).OLEDropMode = 1
'        If Picture1(i).Picture <> 0 Then
'            Picture1(i).PaintPicture Picture1(i).Picture, _
'            0, 0, Picture1(i).ScaleWidth, Picture1(i).ScaleHeight, _
'            0, 0, Picture1(i).Picture.Width / 26.46, _
'            Picture1(i).Picture.Height / 26.46
'        End If
'    Next i
'    Picture1(0).Picture = Picture1(0).Image
End Sub

Private Sub cektbl()
    Strsql = "SELECT * From information_schema.Columns WHERE table_name='tblsystemtraining';"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        qcr = "create table tblsystemtraining (nama_file varchar, addby varchar, id serial not null);"
        M_OBJCONN.execute qcr
    End If
    
    Strsql = "SELECT * From information_schema.Columns WHERE table_name='tblsystemtraining_schedule';"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        qcr = "create table tblsystemtraining_schedule (jam_awal timestamp without time zone, jam_akhir timestamp without time zone, idp integer, ids serial not null);"
        M_OBJCONN.execute qcr
    End If

    Strsql = "SELECT * From information_schema.Columns WHERE table_name='tblsystemtraining_partisipan';"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        qcr = "create table tblsystemtraining_partisipan (agent varchar, ids integer, idme serial not null, f_done integer);"
        M_OBJCONN.execute qcr
    End If

    Strsql = "SELECT * From information_schema.Columns WHERE table_name='tblsystemtraining_partisipan_log';"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        qcr = "create table tblsystemtraining_partisipan_log (agent varchar, ids integer, idme serial not null, f_done integer);"
        M_OBJCONN.execute qcr
    End If
    
End Sub

Private Sub Picture1_DblClick(Index As Integer)
    Dim aaa As String
    
    Set Picture1(Index).Picture = Nothing
    
    On Error GoTo bawah

    CommonDialog1.action = 1
    aaa = CommonDialog1.FileName
    
    Picture1(Index).ScaleMode = 3
    Picture1(Index).AutoRedraw = True
    Picture1(Index).Picture = LoadPicture(aaa)
        Picture1(Index).PaintPicture Picture1(Index).Picture, _
        0, 0, Picture1(Index).ScaleWidth, Picture1(Index).ScaleHeight, _
        0, 0, Picture1(Index).Picture.Width / 26.46, _
        Picture1(Index).Picture.Height / 26.46
    Picture1(Index).Picture = Picture1(Index).Image
    
    Label1(Index).Caption = aaa
    Exit Sub
        
bawah:
    MsgBox "Picture Error"
End Sub

Private Sub Text1_Click()
    If Text1.text <> "" Then
        Strsql = "SELECT * From tblsystemtraining where nama_file = '" & Text1.text & "';"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
        If rs.RecordCount <> 0 Then
            Label8.Caption = cnull(rs!ID)
        End If
    End If
End Sub

Private Sub Text1_DropDown()
    Strsql = "SELECT * From tblsystemtraining;"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    Text1.clear
    
    If rs.RecordCount <> 0 Then
        For i = 1 To rs.RecordCount
            Text1.AddItem rs!nama_file
            rs.MoveNext
        Next i
    End If
End Sub
