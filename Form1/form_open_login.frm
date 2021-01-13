VERSION 5.00
Begin VB.Form form_open_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPEN LOGIN"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   1845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "form_open_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    query = "UPDATE usertbl SET f_status_login=null,last_logout='now()' where userid = '" + Text1.text + "'"
    M_OBJCONN.Execute query
    
    MsgBox "Data Sudah Terupdate"
End Sub
