VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form_report_feedback 
   Caption         =   "Report FeedBack"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14175
   LinkTopic       =   "Form3"
   ScaleHeight     =   8265
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria Report"
      Height          =   2100
      Left            =   -105
      TabIndex        =   0
      Top             =   0
      Width           =   14280
      Begin VB.CommandButton cmbemail 
         Caption         =   "Send To Email"
         Height          =   360
         Left            =   5655
         Picture         =   "form_report_feedback.frx":0000
         TabIndex        =   14
         Top             =   1455
         Width           =   1605
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3225
         TabIndex        =   9
         Top             =   615
         Width           =   2370
      End
      Begin VB.ComboBox cboagentname 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   615
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E87211&
         Caption         =   "Show Phone Number"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   13995
         TabIndex        =   7
         Top             =   -1290
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton SSCommand2 
         BackColor       =   &H00F1E5DB&
         Caption         =   "Batal"
         Height          =   375
         Left            =   5655
         Picture         =   "form_report_feedback.frx":05EE
         TabIndex        =   6
         Top             =   630
         Width           =   1605
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Show"
         Height          =   360
         Left            =   5655
         Picture         =   "form_report_feedback.frx":0C34
         TabIndex        =   5
         Top             =   255
         Width           =   1605
      End
      Begin VB.CommandButton SSCommand1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         Height          =   405
         Left            =   5655
         Picture         =   "form_report_feedback.frx":1222
         TabIndex        =   4
         Top             =   1020
         Width           =   1590
      End
      Begin VB.ComboBox cbocampaign 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   270
         Width           =   4035
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   13275
         TabIndex        =   2
         Top             =   -930
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtlead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13125
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1680
         Width           =   915
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   13755
         Top             =   -570
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xls"
      End
      Begin Crystal.CrystalReport RPT 
         Left            =   13275
         Top             =   -570
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Telesales "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   270
         TabIndex        =   12
         Top             =   615
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Campaign"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Lead"
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
         Height          =   255
         Left            =   12120
         TabIndex        =   10
         Top             =   1695
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6150
      Left            =   0
      TabIndex        =   13
      Top             =   2100
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   10848
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "form_report_feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
