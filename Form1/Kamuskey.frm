VERSION 5.00
Begin VB.Form Kamuskey 
   Caption         =   "Kamus Key"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2790
   LinkTopic       =   "Form5"
   ScaleHeight     =   1140
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "KeyAscii"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "KeyCode"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Kamuskey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Label3.Caption = KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Label4.Caption = KeyAscii
End Sub
