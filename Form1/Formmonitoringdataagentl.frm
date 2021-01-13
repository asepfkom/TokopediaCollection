VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formmonitoringdataagentl 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitoring Data Agent TL"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14745
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14745
   Begin VB.CommandButton Command2 
      Caption         =   "Tracking Agent"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2040
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trade"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   12621
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   7155
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   12621
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "Data per TL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Data di Agent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Formmonitoringdataagentl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "User Id", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "Nama Agent", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "Team", 8 * 120
    ListView1.ColumnHeaders.ADD 4, , "Jumlah Data", 8 * 120
    ListView1.ColumnHeaders.ADD 5, , "Regular", 8 * 120
    ListView1.ColumnHeaders.ADD 6, , "Jumlah Data - Regular", 8 * 120
    
    ListView2.ColumnHeaders.ADD 1, , "TL", 10 * 120
    ListView2.ColumnHeaders.ADD 2, , "Jumlah Data", 20 * 120
    ListView2.ColumnHeaders.ADD 3, , "Jumlah Reguler", 8 * 120
    ListView2.ColumnHeaders.ADD 4, , "Jumlah min Reguler", 8 * 120
    ListView2.ColumnHeaders.ADD 5, , "Maximum Data", 8 * 120
End Sub


Private Sub isi()
    '<28sept2017
    'query = "select usertbl.userid, usertbl.agent ,usertbl.team,count(mgm.custid) jml_data from mgm,usertbl where mgm.agent = usertbl.userid and (mgm.agent ilike 'D%' or mgm.agent ilike 'TL%') and mgm.agent not ilike '%deceas%' "
    '-----------
    
    '28sept2017
'    query = " select usertbl.userid,usertbl.agent,usertbl.team,count(mgm.custid) jml_data from ("
'    query = query + " select userid,agent,team from usertbl where aktif = 0 and (userid ilike 'D%' or userid ilike 'TL%') and userid not ilike '%decease%'  "
'    If MDIForm1.Text2.text = "TeamLeader" Then
'        query = query + " and team = '" + MDIForm1.Text1.text + "' "
'    End If
'    query = query + " ) usertbl left join mgm on usertbl.userid = mgm.agent "
'
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open query + " group by 1,2,3 order by 3,1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    '28nov2017
    query = " select a.*, b.reguler, jml_data - reguler as ""jml-reg"" from ( "
    query = query + " select usertbl.userid,usertbl.agent,usertbl.team,count(mgm.custid) jml_data from ( select userid,agent,team from usertbl where aktif = 0 and (userid ilike 'D%' or userid ilike 'TL%') and userid not ilike '%decease%'   "
        If MDIForm1.Text2.text = "TeamLeader" Then
            query = query + " and team = '" + MDIForm1.Text1.text + "' "
        End If
    query = query + " ) usertbl left join mgm on usertbl.userid = mgm.agent  group by 1,2,3 order by 3,1) a"
    query = query + " Left Join"
    query = query + " (select agent, count(agent) reguler from mgm b, ("
    query = query + " select custid from ("
    query = query + " select custid, count(custid) regular from ("
    query = query + " select custid, to_char(paydate, 'yyyy-mm'), sum(payment)  from tbllunas where to_char(paydate, 'yyyy-mm') >= to_char(now() - interval '3 month', 'yyyy-mm')  group by 1,2"
    query = query + " ) a group by 1"
    query = query + " order by 2 desc ) abc where regular > 1"
    query = query + " ) a where a.custid = b.custid group by 1) b on a.userid = b.agent order by 3,1"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    cmdsql = "select jml from dataperagent"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    abc = rs!jml
    
    While Not M_Objrs.EOF
         Set listItem = ListView1.ListItems.ADD(, , M_Objrs("userid"))
             listItem.SubItems(1) = M_Objrs("AGENT")
             listItem.SubItems(2) = M_Objrs("team")
             listItem.SubItems(3) = IIf(IsNull(M_Objrs("jml_data")), "", M_Objrs("jml_data"))
             listItem.SubItems(4) = cnull(M_Objrs("reguler"))
             listItem.SubItems(5) = IIf(IsNull(M_Objrs("jml-reg")), "", M_Objrs("jml-reg"))
             
         If M_Objrs("jml-reg") > abc Then
             listItem.ForeColor = vbRed
             listItem.ListSubItems(1).ForeColor = vbRed
             listItem.ListSubItems(2).ForeColor = vbRed
             listItem.ListSubItems(3).ForeColor = vbRed
         End If
        M_Objrs.MoveNext
    Wend
        M_Objrs.Close
        Set M_Objrs = Nothing
        
'    query = "select team, agent, sum(jml_data) jml_team, count(userid) * " & abc
'    query = query + " max_data_team  from ( "
'    query = query + " select usertbl.userid, usertbl.agent ,usertbl.team,count(mgm.custid) jml_data from mgm,usertbl where mgm.agent = usertbl.userid and (mgm.agent ilike 'D%' or mgm.agent ilike 'TL%') and mgm.agent not ilike '%deceas%' group by 1,2,3 order by 3,1 ) a group by 1,2 order by 1"
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    
    '<28sept2017
'    query = " select ab.*, bc.jml_pertl, ac.count * " & abc & " jml_max "
'    query = query + " from ("
'    query = query + " select a.userid,a.agent,b.count jml_agent from ("
'    query = query + " select userid,agent from usertbl where userid ilike 'TL%' and aktif = 0 order by 1) a, ("
'    query = query + " select count(userid), team from usertbl where (userid ilike 'D%' or userid ilike 'TL%') and team ilike 'TL%' and aktif = 0 group by 2 order by 2,1 ) b where a.userid = b.team) ab,"
'    query = query + " (select count(a.agent) jml_pertl, team from ("
'    query = query + " select agent from mgm ) a, (select userid, team from usertbl where (userid ilike 'D%' or userid ilike 'TL%') and team ilike 'TL%' and aktif = 0) b"
'    query = query + " where a.agent=b.userid group by 2) bc, (select count(userid),team from usertbl where userid ilike 'D%' and team ilike 'TL%' and aktif = 0 group by 2 order by 2) ac where ab.userid = bc.team and bc.team = ac.team"
    '-------------------------------
    
    '28sept2017
'    'query = " select *, jml_agent * " & abc & " jml_max from ( " & vbCrLf
'    query = " select userid,agent,coalesce(jml_agent,0) jml_agent,team,coalesce(jml_pertl,0) jml_pertl, coalesce(jml_agent * " & abc & " ,0) jml_max from ( " & vbCrLf
'    query = query + " select team.userid, team.agent, agent.jml_agent from ( " & vbCrLf
''    query = query + " select userid,agent from usertbl where userid ilike 'TL%' and aktif = 0 order by 1 ) team, " & vbCrLf
'    query = query + " select userid,agent from usertbl where userid ilike 'TL%' and aktif = 0 "
'    If MDIForm1.Text2.text = "TeamLeader" Then
'        query = query + " and team = '" + MDIForm1.Text1.text + "' " & vbCrLf
'    End If
'    query = query + " order by 1 ) team left join " & vbCrLf
''    query = query + " (select count(userid) jml_agent, team from usertbl where (userid ilike 'D%' or userid ilike 'TL%') and team ilike 'TL%' and aktif = 0 group by 2 order by 2,1) agent " & vbCrLf
''    query = query + " where team.userid = agent.team order by 1 ) a left join " & vbCrLf
'    query = query + "  (select count(userid) jml_agent, team from usertbl where userid ilike 'D%' and team ilike 'TL%' and aktif = 0 group by 2 order by 2,1) agent " & vbCrLf
'    query = query + "  on team.userid = agent.team order by 1) a left join " & vbCrLf
'    query = query + " (select team, count(agent) jml_pertl from ( " & vbCrLf
'    query = query + " select agent from mgm) mgm, " & vbCrLf
'    query = query + " (select userid, team from usertbl where (userid ilike 'D%' or userid ilike 'TL%') and team ilike 'TL%' and aktif = 0 order by 1) usertbl where mgm.agent = usertbl.userid " & vbCrLf
'    query = query + " group by 1 order by 1) b on a.userid = b.team "
    
    query = " select team, sum(jml_data) as jumlah_data, sum(reguler) as jumlah_reguler, sum(""jml-reg"") as jumlahminreg, (count(team)-1) * " & abc & " as dataseharusnya from ("
    query = query + " select a.*, b.reguler, jml_data - coalesce(reguler,0) as ""jml-reg"" from ("
    query = query + " select usertbl.userid,usertbl.agent,usertbl.team,count(mgm.custid) jml_data from ( select userid,agent,team from usertbl where aktif = 0 and (userid ilike 'D%' or userid ilike 'TL%') and userid not ilike '%decease%'   ) usertbl left join mgm on usertbl.userid = mgm.agent  group by 1,2,3 order by 3,1) a"
    query = query + " Left Join"
    query = query + " (select agent, count(agent) reguler from mgm b, ("
    query = query + " select custid from ("
    query = query + " select custid, count(custid) regular from ("
    query = query + " select custid, to_char(paydate, 'yyyy-mm'), sum(payment)  from tbllunas where to_char(paydate, 'yyyy-mm') >= to_char(now() - interval '3 month', 'yyyy-mm')  group by 1,2"
    query = query + " ) a group by 1"
    query = query + " order by 2 desc ) abc where regular > 1"
    query = query + " ) a where a.custid = b.custid group by 1) b on a.userid = b.agent order by 3,1) a where team ilike 'TL%' "
        If MDIForm1.Text2.text = "TeamLeader" Then
            query = query + " and team = '" + MDIForm1.Text1.text + "' "
        End If
    query = query + " group by 1"
    
    
    rs1.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs1.EOF
         Set listItem = ListView2.ListItems.ADD(, , rs1("team"))
             listItem.SubItems(1) = rs1("jumlah_data")
             listItem.SubItems(2) = IIf(IsNull(rs1("jumlah_reguler")), "", rs1("jumlah_reguler"))
             listItem.SubItems(3) = IIf(IsNull(rs1("jumlahminreg")), "", rs1("jumlahminreg"))
             listItem.SubItems(4) = IIf(IsNull(rs1("dataseharusnya")), "", rs1("dataseharusnya"))
             
         If rs1("jumlahminreg") > rs1("dataseharusnya") Then
             listItem.ForeColor = vbRed
             listItem.ListSubItems(1).ForeColor = vbRed
             listItem.ListSubItems(2).ForeColor = vbRed
             listItem.ListSubItems(3).ForeColor = vbRed
         End If
        rs1.MoveNext
    Wend
        rs1.Close
        Set rs1 = Nothing
    
End Sub

Private Sub Command1_Click()
    form_trade.Show
End Sub

Private Sub Command2_Click()
    form_tracking_agent.Show
End Sub

Private Sub Form_Load()
    Call header
    Call isi
    If MDIForm1.Text2.text = "Supervisor" Or MDIForm1.Text2.text = "Administrator" Or MDIForm1.Text2.text Like "*Manager*" Then
        Command1.Visible = True
        Command1.Enabled = True
        Command2.Visible = True
        Command2.Enabled = True
    ElseIf Left(MDIForm1.Text2.text, 2) = "AM" Then
        Command2.Visible = True
        Command2.Enabled = True
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
End Sub
