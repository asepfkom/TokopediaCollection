VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMTERIMAMSG 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6240
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   120
      TabIndex        =   8
      Top             =   5220
      Width           =   4335
      Begin VB.CommandButton CmdRequestNumber 
         Caption         =   "&Request Number"
         Height          =   435
         Left            =   1620
         TabIndex        =   10
         Top             =   300
         Width           =   1515
      End
      Begin VB.CommandButton CmdListReqPTP 
         Caption         =   "&List Req.PTP"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   75
      TabIndex        =   7
      Top             =   4695
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   195
      Width           =   3525
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   4020
      Left            =   0
      TabIndex        =   5
      Top             =   525
      Width           =   4500
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3840
         Left            =   30
         TabIndex        =   4
         Top             =   135
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   6773
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"FRMTERIMAMSG.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   3645
      TabIndex        =   2
      Top             =   4695
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Teruskan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2655
      TabIndex        =   1
      Top             =   4695
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Balas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   4695
      Width           =   780
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Dari :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   195
      TabIndex        =   6
      Top             =   255
      Width           =   465
   End
End
Attribute VB_Name = "FRMTERIMAMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdListReqPTP_Click()
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    FrmListRequestPTP.Show vbModal
    Me.Hide
End Sub

Private Sub CmdRequestNumber_Click()
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    Me.Hide
    FrmListReqTlp.Show vbModal
    
End Sub

Private Sub Command1_Click(Index As Integer)
Dim cmdsql As String

Select Case Index
    Case 0
        FRMSENDMSG.Text1.Text = Text1.Text + ";"
        FRMSENDMSG.Command2.Enabled = False
        Unload Me
        'FRMSENDMSG.Show vbModal
        FRMSENDMSG.Show vbModal
    Case 1
        FRMSENDMSG.RichTextBox1.Text = RichTextBox1.Text
        Unload Me
        'FRMSENDMSG.Show vbModal
        FRMSENDMSG.Show vbModal
    Case 2
        Unload Me
    Case 3
      '  CMDSQL = "Insert into         "
End Select
End Sub

Private Sub Form_Load()
    Call BringWindowToTop(Me.hwnd)
End Sub
