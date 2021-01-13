VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form show_logreupload 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "History ReUpload"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7980
   LinkTopic       =   "Form3"
   ScaleHeight     =   4545
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "History ReUpload"
      Height          =   4590
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   7980
      Begin MSComctlLib.ListView listview1 
         Height          =   4335
         Index           =   1
         Left            =   45
         TabIndex        =   1
         Top             =   195
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6120
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "show_logreupload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim listview1(1) As ListItem
   Call HEADER_LOG_REUPLOAD
   Call Show_Mapping_history_reupload
End Sub

Private Sub HEADER_LOG_REUPLOAD()
    listview1(1).ColumnHeaders.ADD 1, , "No", 5 * TXT
    listview1(1).ColumnHeaders.ADD 2, , "Tgl Reupload", 15 * TXT
    listview1(1).ColumnHeaders.ADD 3, , "Recsource", 10 * TXT
    listview1(1).ColumnHeaders.ADD 4, , "Agent", 10 * TXT
    listview1(1).ColumnHeaders.ADD 5, , "Dispo", 10 * TXT
    listview1(1).ColumnHeaders.ADD 6, , "Jumlah", 10 * TXT
    listview1(1).ColumnHeaders.ADD 7, , "User Input", 12 * TXT
    
End Sub
Private Function Show_Mapping_history_reupload()
Dim rs As New ADODB.Recordset
Dim ListItem As ListItem
     rs.CursorLocation = adUseClient
     rs.Open "select tgl_reupload,recsource,agent,f_cek_new,jumlah,user_input from reupload_log " & _
     " where date(tgl_reupload) = date(now())", M_OBJCONN, adOpenDynamic, adLockOptimistic
     listview1(1).ListItems.clear
     While Not rs.EOF
     Me.Refresh
         Set ListItem = listview1(1).ListItems.ADD(, , rs.Bookmark)
            ListItem.SubItems(1) = IIf(IsNull(rs("tgl_reupload")), "", rs("tgl_reupload"))
            ListItem.SubItems(2) = IIf(IsNull(rs("recsource")), "", rs("recsource"))
            ListItem.SubItems(3) = IIf(IsNull(rs("agent")), "", rs("agent"))
            ListItem.SubItems(4) = IIf(IsNull(rs("f_cek_new")), "", rs("f_cek_new"))
            ListItem.SubItems(5) = IIf(IsNull(rs("jumlah")), "", rs("jumlah"))
            ListItem.SubItems(6) = IIf(IsNull(rs("user_input")), "", rs("user_input"))
            rs.MoveNext
        Wend

End Function

