VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame18 
      BackColor       =   &H00ABE18E&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin MSComctlLib.ListView LstReserve 
         Height          =   1005
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1773
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   10147522
         BorderStyle     =   1
         Appearance      =   0
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
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   3
         Left            =   2490
         TabIndex        =   2
         Top             =   255
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "Form5.frx":0000
         AutoSize        =   1
         Alignment       =   8
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H003F9E0C&
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2490
         TabIndex        =   3
         Top             =   825
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
