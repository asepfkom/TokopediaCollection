VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_old_remarks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OLD REMARKS"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12330
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_old_remarks.frx":0000
      Left            =   3240
      List            =   "frm_old_remarks.frx":000D
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   7320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   7080
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   12488
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7440
      Width           =   375
   End
   Begin VB.Label lblCustId 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   7320
      Width           =   2655
   End
End
Attribute VB_Name = "frm_old_remarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn_other      As ADODB.Connection
Private conn_sql        As String
Private rs              As ADODB.Recordset
Public strIdLbl         As String

Private Function get_table_name() As String
    
    If Combo1.Text = "" Then
        get_table_name = "backup_mgm_hst_2016"
    Else
        get_table_name = "backup_mgm_hst_" & Trim(Combo1.Text)
    End If
End Function

Private Sub Combo1_Click()
    Call load_old_remarks
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim cmd_sql     As String
    Dim sTime_Hst   As String

    Call koneksi
    Call HEADER_HISTORY
    'Combo1.ListIndex = 1
    Timer1.Enabled = True
        
    lblCustId.Caption = strIdLbl
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    'Set conn_other = Nothing
End Sub

Private Sub load_old_remarks()
    cmd_sql = "SELECT id,custid,tgl,agent,coalesce(hst,'') as hst,KodeDS,KdComplaint,f_cek ,statuscall ,ststelpwith,user_log,stop_time,f_special"
    cmd_sql = cmd_sql + " FROM " + get_table_name
    cmd_sql = cmd_sql + " WHERE custid = '" + Trim(lblCustId.Caption) + "'"
    cmd_sql = cmd_sql + " ORDER BY tgl DESC "

    If rs.state = 1 Then rs.Close
    rs.Open cmd_sql
    listview1.ListItems.CLEAR
    While Not rs.EOF
        'Set listitem = ListView1(1).ListItems.ADD(, , Left(rs("TGL"), 4) & "/" & Mid(rs("TGL"), 5, 2) & "/" & IIf(IsNull(rs("TGL")), "", Mid(rs("TGL"), 7, 2)) & " " & IIf(IsNull(rs("TGL")), "", Mid(rs("TGL"), 9, 2)) & ":" & Right(rs("TGL"), 2))
        sTime_Hst = ""
        If IIf(IsNull(rs("tgl")), "", rs!TGL) <> "" Then
            'sTime_Hst = Format(IIf(IsNull(rs("TGL")), "", rs!TGL), "mm-dd-yyyy hh:mm:ss") & Format(IIf(IsNull(rs("stop_time")), "", rs!stop_time), " - hh:mm:ss")
           sTime_Hst = Format(IIf(IsNull(rs("tgl")), "", rs!TGL), "mm-dd-yyyy hh:mm:ss")
        End If
        
        Set listItem = listview1.ListItems.ADD(, , sTime_Hst)
        listItem.SubItems(1) = IIf(IsNull(rs("hst")), "", rs("hst"))
        listItem.SubItems(2) = IIf(IsNull(rs("user_log")), "", rs("user_log"))
        listItem.SubItems(3) = IIf(IsNull(rs("agent")), "", rs("agent"))
        listItem.SubItems(4) = IIf(IsNull(rs("KodeDs")), "", rs("KodeDs"))
        listItem.SubItems(5) = IIf(IsNull(rs("statuscall")), "", rs("statuscall"))
        listItem.SubItems(6) = IIf(IsNull(rs("ststelpwith")), "", rs("ststelpwith"))
        listItem.SubItems(7) = IIf(IsNull(rs("id")), "", rs("id"))
        'listitem.SubItems(4) = IIf(IsNull(rs("f_cek")), "", rs("f_cek"))
        'Data Special 'jejaktian 18032016
        If IIf(IsNull(rs("f_special")), 0, rs("f_special")) = "1" Then
            For K = 1 To 7
                listItem.ListSubItems(K).ForeColor = vbRed
                listItem.ListSubItems(K).Bold = True
            Next K
        End If
        ' ------------------------------------------
        rs.MoveNext
    Wend
End Sub

Private Sub Timer1_Timer()
    Call load_old_remarks
    Timer1.Enabled = False
End Sub

Private Sub koneksi()
'    conn_sql = "Provider=MSDASQL.1;Persist Security Info=False;User ID=program_db_access;PWD=program_db_access123;Data Source=RITCARD_POSTGRE"
'    conn_other.Open conn_sql
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = M_OBJCONN
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
End Sub

Private Sub HEADER_HISTORY()
    listview1.ColumnHeaders.ADD 1, , "Tanggal(mm-dd-yyyy)", 10 * TXT
    listview1.ColumnHeaders.ADD 2, , "History", 80 * TXT
    listview1.ColumnHeaders.ADD 3, , "User Log", 10 * TXT
    listview1.ColumnHeaders.ADD 4, , "Handle By", 10 * TXT
    listview1.ColumnHeaders.ADD 5, , "Sts Account", 10 * TXT
    listview1.ColumnHeaders.ADD 6, , "Sts Call", 10 * TXT
    listview1.ColumnHeaders.ADD 7, , "Sts Telp With", 25 * TXT
    listview1.ColumnHeaders.ADD 8, , "Id", 25 * TXT
End Sub

