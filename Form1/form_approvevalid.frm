VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form form_approvevalid 
   Caption         =   "Approval Valid Phone"
   ClientHeight    =   5925
   ClientLeft      =   525
   ClientTop       =   765
   ClientWidth     =   7980
   LinkTopic       =   "Form5"
   ScaleHeight     =   5925
   ScaleWidth      =   7980
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete From Existing "
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Check"
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Nomor Valid"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CustID :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reject"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Log"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   4560
      Width           =   1300
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "form_approvevalid.frx":0000
      Left            =   4440
      List            =   "form_approvevalid.frx":000A
      TabIndex        =   9
      Top             =   80
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Approve"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   75
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   80
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3870
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6826
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
      Appearance      =   1
      MousePointer    =   1
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   3870
      Left            =   8280
      TabIndex        =   11
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6826
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
      Appearance      =   1
      MousePointer    =   1
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
      Caption         =   "0"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "History"
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
      Left            =   14280
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Be"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah : 0"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "TL"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Agent"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "form_approvevalid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub isicombo()
    Combo1.CLEAR
    Combo2.CLEAR
    If MDIForm1.Text2.text = "TeamLeader" Or Combo3.text = "TeamLeader" Then
        
        
        q = "select distinct agent from tblvalidtotl "
        If MDIForm1.Text2.text = "TeamLeader" Then
            Combo2.text = MDIForm1.Text1.text
            q = q + "where spv = '" + Combo2.text + "'"
        End If
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q + " order by 1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
        While Not r.EOF
            Combo1.AddItem r!agent
            r.MoveNext
        Wend
                
        Set q = Nothing
        
        q = "select distinct spv from tblvalidtotl "
        If MDIForm1.Text2.text = "TeamLeader" Then
            q = q + "where spv = '" + Combo2.text + "'"
        End If
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q + " order by 1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not r.EOF
            Combo2.AddItem r!spv
            r.MoveNext
        Wend
        
        Set q = Nothing
    ElseIf MDIForm1.Text2.text = "Supervisor" Or MDIForm1.Text2.text = "Administrator" Or Combo3.text = "Supervisor" Then
        q = "select distinct agent from tblvalidtospv"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q + " order by 1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not r.EOF
            Combo1.AddItem r!agent
            r.MoveNext
        Wend
        
        Set q = Nothing
        
        q = "select distinct spv from tblvalidtospv"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q + " order by 1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not r.EOF
            Combo2.AddItem r!spv
            r.MoveNext
        Wend
        
        Set q = Nothing
        
    End If
End Sub

Private Sub isilist()
        ListView1.ColumnHeaders.CLEAR
        ListView1.ListItems.CLEAR
    If MDIForm1.Text2.text = "TeamLeader" Or Combo3.text = "TeamLeader" Then
        ListView1.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
        ListView1.ColumnHeaders.ADD 2, , "Nomor Valid", 20 * 120
        ListView1.ColumnHeaders.ADD 3, , "Agent", 20 * 120
        ListView1.ColumnHeaders.ADD 4, , "TL", 10 * 120
        ListView1.ColumnHeaders.ADD 5, , "Tanggal Request Agent", 20 * 120
        
        q = "select * from tblvalidtotl "
        If MDIForm1.Text2.text = "TeamLeader" Or Combo2.text <> "" Then
            q = q + "where spv = '" + Combo2.text + "' "
        End If
        If Combo1.text <> "" Then
            q = q + " and agent = '" + Combo1.text + "' "
        End If
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q + "order by tanggalreq", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not r.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(r("custid")))
             listItem.SubItems(1) = cnull(r("nomorvalid"))
             listItem.SubItems(2) = cnull(r("agent"))
             listItem.SubItems(3) = cnull(r("spv"))
             listItem.SubItems(4) = cnull(r("tanggalreq"))
        r.MoveNext
        Wend
                
        Label2.Caption = "Jumlah Data :" & r.RecordCount
        Set r = Nothing
        
    ElseIf MDIForm1.Text2.text = "Supervisor" Or MDIForm1.Text2.text = "Administrator" Or Combo3.text = "Supervisor" Then
        ListView1.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
        ListView1.ColumnHeaders.ADD 2, , "Nomor Valid", 20 * 120
        ListView1.ColumnHeaders.ADD 3, , "Agent", 10 * 120
        ListView1.ColumnHeaders.ADD 4, , "TL", 10 * 120
        ListView1.ColumnHeaders.ADD 5, , "Tanggal Request Agent", 20 * 120
        ListView1.ColumnHeaders.ADD 6, , "Tanggal Request TL", 20 * 120
        
        q = "select * from tblvalidtospv "
        If MDIForm1.Text2.text = "TeamLeader" Or Combo2.text <> "" Then
            q = q + "where spv = '" + Combo2.text + "' "
        End If
        If Combo1.text <> "" Then
            q = q + " and agent = '" + Combo1.text + "' "
        End If
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q + "order by tanggalreq", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not r.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(r("custid")))
             listItem.SubItems(1) = cnull(r("nomorvalid"))
             listItem.SubItems(2) = cnull(r("agent"))
             listItem.SubItems(3) = cnull(r("spv"))
             listItem.SubItems(4) = cnull(r("tanggalreq"))
             listItem.SubItems(5) = cnull(r("tanggalapptl"))
        r.MoveNext
        Wend
        Label2.Caption = "Jumlah Data :" & r.RecordCount
        Set r = Nothing
    End If
End Sub

Private Sub Command1_Click()
    isicombo
    Call isilist
End Sub

Private Sub Command2_Click()
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Search Data terlebih dahulu"
        Exit Sub
    End If
    For K = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    Call insertapp
End Sub

Private Sub insertapp()
    If MDIForm1.Text2.text = "TeamLeader" Or Combo3.text = "TeamLeader" Then
        For w = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(w).Checked = True Then
                CustId = ListView1.ListItems(w).text
                nomor = ListView1.ListItems(w).ListSubItems(1)
                agent = ListView1.ListItems(w).ListSubItems(2)
                spv = ListView1.ListItems(w).ListSubItems(3)
                Tanggal = ListView1.ListItems(w).ListSubItems(4)
                
                query = "INSERT INTO tblvalidtotllog values ('" + CustId + "','" + nomor + "','" + agent + "','" + spv + "', '" + Tanggal + "') ;" & vbCrLf
                query = query + "INSERT INTO tblvalidtospv values ('" + CustId + "','" + nomor + "','" + agent + "','" + spv + "', '" + Tanggal + "',now()) ;" & vbCrLf
                query = query + "Delete from tblvalidtotl where custid = '" + CustId + "';"
                M_OBJCONN.Execute query
            End If
        Next w
        MsgBox "Data masuk ke tahap Approve SPV"
    ElseIf MDIForm1.Text2.text = "Supervisor" Or MDIForm1.Text2.text = "Administrator" Or Combo3.text = "Supervisor" Then
        For w = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(w).Checked = True Then
                CustId = ListView1.ListItems(w).text
                nomor = ListView1.ListItems(w).ListSubItems(1)
                agent = ListView1.ListItems(w).ListSubItems(2)
                spv = ListView1.ListItems(w).ListSubItems(3)
                Tanggal = ListView1.ListItems(w).ListSubItems(4)
                tanggaltl = ListView1.ListItems(w).ListSubItems(5)
                approve = MDIForm1.Text1.text
                
                
                query = "INSERT INTO tblvalidtospvlog values ('" + CustId + "','" + nomor + "','" + agent + "','" + spv + "', '" + Tanggal + "', '" + tanggaltl + "', '" + approve + "', now()) ;" & vbCrLf
                query = query + "Update mgm set validsms = '" + nomor + "' where custid = '" + CustId + "';" & vbCrLf
                query = query + "Delete from tblvalidtospv where custid = '" + CustId + "';"
                M_OBJCONN.Execute query
            End If
        Next w
        MsgBox "Data Approved"
    End If
    Command1_Click
End Sub

Private Sub deleteapp()
    If MDIForm1.Text2.text = "TeamLeader" Or Combo3.text = "TeamLeader" Then
        For w = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(w).Checked = True Then
                CustId = ListView1.ListItems(w).text
                nomor = ListView1.ListItems(w).ListSubItems(1)
                agent = ListView1.ListItems(w).ListSubItems(2)
                spv = ListView1.ListItems(w).ListSubItems(3)
                Tanggal = ListView1.ListItems(w).ListSubItems(4)
                
                query = "INSERT INTO tblvalidtotllog values ('" + CustId + "','" + nomor + "','" + agent + "','" + spv + "', '" + Tanggal + "', 'REJECT') ;" & vbCrLf
                query = query + "Delete from tblvalidtotl where custid = '" + CustId + "';"
                M_OBJCONN.Execute query
            End If
        Next w
        MsgBox "Data masuk ke tahap Approve SPV"
    ElseIf MDIForm1.Text2.text = "Supervisor" Or MDIForm1.Text2.text = "Administrator" Or Combo3.text = "Supervisor" Then
        For w = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(w).Checked = True Then
                CustId = ListView1.ListItems(w).text
                nomor = ListView1.ListItems(w).ListSubItems(1)
                agent = ListView1.ListItems(w).ListSubItems(2)
                spv = ListView1.ListItems(w).ListSubItems(3)
                Tanggal = ListView1.ListItems(w).ListSubItems(4)
                tanggaltl = ListView1.ListItems(w).ListSubItems(5)
                approve = MDIForm1.Text1.text
                
                
                query = "INSERT INTO tblvalidtospvlog values ('" + CustId + "','" + nomor + "','" + agent + "','" + spv + "', '" + Tanggal + "', '" + tanggaltl + "', '" + approve + "', now(),'REJECT') ;" & vbCrLf
                query = query + "Delete from tblvalidtospv where custid = '" + CustId + "';"
                M_OBJCONN.Execute query
            End If
        Next w
        MsgBox "Data Rejected"
    End If
    Command1_Click
End Sub


Private Sub Command3_Click()
    Call isilog
    If Me.Width = 8100 Then
        Command3.Caption = "Hide Log"
        Me.Width = 16200
    Else
        Command3.Caption = "Show Log"
        Me.Width = 8100
    End If
End Sub

Private Sub Command4_Click()
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Search Data terlebih dahulu"
        Exit Sub
    End If
    For K = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    Call deleteapp

End Sub

Private Sub Command5_Click()
    If Text1.text = "" Then
        MsgBox "Harap Isi Custid yang ingin di delete"
        Exit Sub
    End If
    
    q = "select * from mgm where custid = '" + Text1.text + "' "
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseClient
    r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If r.RecordCount > 0 Then
        MsgBox "Data Ditemukan"
        Command6.Enabled = True
        If cnull(r!validsms) <> "" Then
            Label6.Caption = cnull(r!validsms)
        Else
            Label6.Caption = "Nomor Valid Kosong"
        End If
    Else
        MsgBox "Data Tidak Ditemukan"
        Exit Sub
    End If
End Sub

Private Sub Command6_Click()
    q = "update mgm set validsms = '' where custid = '" + Text1.text + "' "
    M_OBJCONN.Execute q
    
    MsgBox "Nomor Valid berhasil dihapus"
    Text1.text = ""
    Label6.Caption = "Nomor Valid"
    Command6.Enabled = False
End Sub

Private Sub Form_Load()
    If MDIForm1.Text2.text <> "TeamLeader" Then
        Label1(1).Visible = True
        Label1(2).Visible = True
        Frame1.Visible = True
        
        Combo2.Visible = True
        Combo3.Visible = True
    End If
    
    Label4.Caption = 0
    
    Me.Width = 8100
    
    Call isicombo
End Sub


Private Sub isilog()
    ListView2.ColumnHeaders.CLEAR
    ListView2.ListItems.CLEAR
    
    ListView2.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
    ListView2.ColumnHeaders.ADD 2, , "Nomor Valid", 20 * 120
    ListView2.ColumnHeaders.ADD 3, , "Agent", 10 * 120
    ListView2.ColumnHeaders.ADD 4, , "TL", 10 * 120
    ListView2.ColumnHeaders.ADD 5, , "Tanggal Request Agent", 10 * 120
    ListView2.ColumnHeaders.ADD 6, , "Tanggal Request TL", 10 * 120
    ListView2.ColumnHeaders.ADD 7, , "Approve By", 10 * 120
    ListView2.ColumnHeaders.ADD 8, , "Tanggal Approve", 10 * 120

    q = "select * from tblvalidtospvlog order by tanggalappspv desc limit 1000"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseClient
    r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not r.EOF
        Set listItem = ListView2.ListItems.ADD(, , cnull(r("custid")))
             listItem.SubItems(1) = cnull(r("nomorvalid"))
             listItem.SubItems(2) = cnull(r("agent"))
             listItem.SubItems(3) = cnull(r("spv"))
             listItem.SubItems(4) = cnull(r("tanggalreq"))
             listItem.SubItems(5) = cnull(r("tanggalapptl"))
             listItem.SubItems(6) = cnull(r("approveby"))
             listItem.SubItems(7) = cnull(r("tanggalappspv"))
        r.MoveNext
    Wend
    Set r = Nothing
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = ListView1.SelectedItem.text
        form_approvevalid.Hide
        VIEW_MGMDATA.Show
        VIEW_MGMDATA.Command1(0).SetFocus
        Sendkeys "{Enter}"
        WaitSecs 0.01
        VIEW_MGMDATA.LstVwSearchMgm.SetFocus
        Sendkeys "{Enter}"
        Label4.Caption = 1
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub Text1_Change()
    Command6.Enabled = False
    Label6.Caption = "Nomor Valid"
End Sub
