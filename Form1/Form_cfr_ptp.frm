VERSION 5.00
Begin VB.Form Form_cfr_ptp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   1620
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Confirm PTP"
      Height          =   1590
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   4575
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "X"
         Height          =   315
         Left            =   4215
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   90
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Update"
         Height          =   375
         Left            =   300
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   945
         Width           =   3915
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form_cfr_ptp.frx":0000
         Left            =   1560
         List            =   "Form_cfr_ptp.frx":000A
         TabIndex        =   1
         Text            =   "BP - Broken Promise"
         Top             =   420
         Width           =   2670
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Status PTP"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   300
         TabIndex        =   2
         Top             =   405
         Width           =   1800
      End
   End
End
Attribute VB_Name = "Form_cfr_ptp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    M_OBJCONN.execute "update mgm set status_ptp = left('" & Combo1.text & "',2) where custid = '" & FrmCC_Colection.TxtCustid.text & "'"
    MsgBox "DONE", vbOKOnly + vbInformation, "Infomasi"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
