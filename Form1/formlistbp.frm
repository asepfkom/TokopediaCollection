VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formlistbp 
   Caption         =   "List BP"
   ClientHeight    =   6015
   ClientLeft      =   180
   ClientTop       =   645
   ClientWidth     =   14655
   LinkTopic       =   "Form5"
   ScaleHeight     =   6015
   ScaleWidth      =   14655
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "List"
      Height          =   5415
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   14655
      Begin VB.CommandButton Command6 
         BackColor       =   &H000000C0&
         Caption         =   "Send To DC"
         Height          =   375
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Fix Data"
         Height          =   375
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check All"
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "Export"
         Height          =   375
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Send To TL"
         Height          =   375
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000A&
         Caption         =   "Log"
         Height          =   375
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4770
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   4770
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4470
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   7885
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
      Begin VB.Label Label1 
         Caption         =   "AMT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "ACC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   4800
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Filter"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL12"
         Height          =   255
         Index           =   12
         Left            =   7440
         MaskColor       =   &H008080FF&
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL11"
         Height          =   255
         Index           =   11
         Left            =   6840
         MaskColor       =   &H008080FF&
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL10"
         Height          =   255
         Index           =   10
         Left            =   6240
         MaskColor       =   &H008080FF&
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL9"
         Height          =   255
         Index           =   9
         Left            =   5640
         MaskColor       =   &H008080FF&
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL8"
         Height          =   255
         Index           =   8
         Left            =   5040
         MaskColor       =   &H008080FF&
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL7"
         Height          =   255
         Index           =   7
         Left            =   4440
         MaskColor       =   &H008080FF&
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL6"
         Height          =   255
         Index           =   6
         Left            =   3840
         MaskColor       =   &H008080FF&
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL5"
         Height          =   255
         Index           =   5
         Left            =   3240
         MaskColor       =   &H008080FF&
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL4"
         Height          =   255
         Index           =   4
         Left            =   2640
         MaskColor       =   &H008080FF&
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL3"
         Height          =   255
         Index           =   3
         Left            =   2040
         MaskColor       =   &H008080FF&
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL2"
         Height          =   255
         Index           =   2
         Left            =   1440
         MaskColor       =   &H008080FF&
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TL1"
         Height          =   255
         Index           =   1
         Left            =   840
         MaskColor       =   &H008080FF&
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ALL"
         Height          =   255
         Index           =   0
         Left            =   240
         MaskColor       =   &H008080FF&
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "formlistbp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Dim r As Integer
        
    If Check1.Value = vbChecked Then
        If ListView1.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = True
        Next r
    Else
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = False
        Next r
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    isilistbtn (Index)
End Sub

Private Sub headerlist()
    ListView1.ColumnHeaders.clear
    ListView1.Checkboxes = True
    ListView1.ColumnHeaders.ADD 1, , "Custid", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "PTP Date", 20 * 120
    ListView1.ColumnHeaders.ADD 4, , "Amt PTP", 8 * 120
    ListView1.ColumnHeaders.ADD 5, , "Product", 10 * 120
    ListView1.ColumnHeaders.ADD 6, , "BP Date (Last)", 20 * 120
    ListView1.ColumnHeaders.ADD 7, , "DC", 20 * 120
    ListView1.ColumnHeaders.ADD 8, , "TL", 8 * 120
End Sub

Private Sub headerlist_log()
    ListView1.ColumnHeaders.clear
    ListView1.Checkboxes = False
    ListView1.ColumnHeaders.ADD 1, , "Custid", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "PTP Date", 20 * 120
    ListView1.ColumnHeaders.ADD 4, , "Amt PTP", 8 * 120
    ListView1.ColumnHeaders.ADD 5, , "Product", 10 * 120
    ListView1.ColumnHeaders.ADD 6, , "BP Date (Last)", 20 * 120
    ListView1.ColumnHeaders.ADD 7, , "DC", 20 * 120
    ListView1.ColumnHeaders.ADD 8, , "TL", 8 * 120
    ListView1.ColumnHeaders.ADD 9, , "Send By", 8 * 120
    ListView1.ColumnHeaders.ADD 10, , "Send Date", 8 * 120
End Sub

Private Sub isilist()

    ListView1.ListItems.clear
    
'    query = " select a.custid,name,promisedate,promisepay,acc_type,tglbp,agentlama,agent_asli from ( "
'    query = vbCrLf & query & "select a.custid, name, acc_type,agentlama,agent_asli,agent,tgl as tglbp,f_cek_new from mgm a,"
'    query = vbCrLf & query & "(select custid, max(tgl) as tgl from mgm_hst where hst in ('BP-NEW', 'BP_POP') group by 1) b where a.custid = b.custid"
'    query = vbCrLf & query & ") a, ("
'    query = vbCrLf & query & "select a.custid,promisedate,promisepay from tblnegoptp a, ("
'    query = vbCrLf & query & "select custid, max(promisedate) as promise_date from tblnegoptp  group by 1) b where a.custid = b.custid and a.promisedate = b.promise_date"
'    query = vbCrLf & query & ") b where a.custid = b.custid and f_cek_new ilike 'BP%' and agent = 'BP' "
    'name,promisedate,promisepay,acc_type,tglstatus,tglbp,agentlama,agent_asli
'    query = " select a.custid, name,promisedate,promisepay, acc_type,tglstatus,agentlama,agent_asli,agent,f_cek_new from mgm a "
'    query = vbCrLf & query & "join ("
'    query = vbCrLf & query & "select a.custid,promisedate,promisepay from tblnegoptp a where id in"
'    query = vbCrLf & query & "("
'    query = vbCrLf & query & "select id from ( select custid,max(id) as id from tblnegoptp group by 1 ) b"
'    query = vbCrLf & query & "))b on a.custid =b.custid and f_cek_new ilike 'BP%' and agent = 'BP'"
    
    query = " select a.*, case when a.promisedate <= b.tgllunas then 1 else 0 end as warna   from ("
    query = vbCrLf & query & " select a.custid, name,promisedate,promisepay, acc_type,tglstatus,agentlama,agent_asli,agent,f_cek_new from mgm a join (select a.custid,promisedate,promisepay from tblnegoptp a where id in(select id from ( select custid,max(id) as id from tblnegoptp group by 1 ) b))b on a.custid =b.custid and f_cek_new ilike 'BP%' and agent = 'BP'"
    
    If MDIForm1.Text2.text <> "Supervisor" And MDIForm1.Text2.text <> "Manager" Then
    
        query1 = " select * from  usertbl where userid = '" + MDIForm1.Text1.text + "'"
        Set rs1 = New ADODB.Recordset
        rs1.CursorLocation = adUseClient
        rs1.Open query1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        query = vbCrLf & query & " and agent_asli = '" + rs1!TEAM + "'"
    End If
    
    query = vbCrLf & query & " ) a left join vwtbllunaslast b on a.custid = b.custid"
    
    
    
    If MDIForm1.Text2.text <> "Supervisor" And MDIForm1.Text2.text <> "Manager" Then
    
        query1 = " select * from  usertbl where userid = '" + MDIForm1.Text1.text + "'"
        Set rs1 = New ADODB.Recordset
        rs1.CursorLocation = adUseClient
        rs1.Open query1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        query = vbCrLf & query & " and agent_asli = '" + rs1!TEAM + "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    jmlamt = 0
    
    While Not rs.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = Format(cnull(rs("promisedate")), "YYYY-MM-DD")
             listItem.SubItems(3) = cnull(rs("promisepay"))
             jmlamt = jmlamt + cnull(rs("promisepay"))
             listItem.SubItems(4) = cnull(rs("acc_type"))
             listItem.SubItems(5) = Format(cnull(rs("tglstatus")), "YYYY-MM-DD hh:nn:ss")
             listItem.SubItems(6) = cnull(rs("agentlama"))
             listItem.SubItems(7) = cnull(rs("agent_asli"))
             
                If rs("warna") = 1 Then
                    listItem.ListSubItems(1).ForeColor = vbBlue
                    listItem.ListSubItems(2).ForeColor = vbBlue
                    listItem.ListSubItems(3).ForeColor = vbBlue
                    listItem.ListSubItems(4).ForeColor = vbBlue
                    listItem.ListSubItems(5).ForeColor = vbBlue
                    listItem.ListSubItems(6).ForeColor = vbBlue
                    listItem.ListSubItems(7).ForeColor = vbBlue
                    'listItem.ListSubItems(8).ForeColor = vbBlue
                    'listItem.ListSubItems(9).ForeColor = vbBlue
                End If
        rs.MoveNext
    Wend
    Text1.text = rs.RecordCount
    Text2.text = Format(jmlamt, "##,###")
End Sub

Private Sub isilistbtn(Index As Integer)

    ListView1.ListItems.clear
    
'    query = " select a.custid,name,promisedate,promisepay,acc_type,tglbp,agentlama,agent_asli from ( "
'    query = vbCrLf & query & "select a.custid, name, acc_type,agentlama,agent_asli,agent,tgl as tglbp,f_cek_new from mgm a,"
'    query = vbCrLf & query & "(select custid, max(tgl) as tgl from mgm_hst where hst in ('BP-NEW', 'BP_POP') group by 1) b where a.custid = b.custid"
'    query = vbCrLf & query & ") a, ("
'    query = vbCrLf & query & "select a.custid,promisedate,promisepay from tblnegoptp a, ("
'    query = vbCrLf & query & "select custid, max(promisedate) as promise_date from tblnegoptp  group by 1) b where a.custid = b.custid and a.promisedate = b.promise_date"
'    query = vbCrLf & query & ") b where a.custid = b.custid and f_cek_new ilike 'BP%' and agent = 'BP' "
    
'    query = " select a.custid, name,promisedate,promisepay, acc_type,tglstatus,agentlama,agent_asli,agent,f_cek_new from mgm a "
'    query = vbCrLf & query & "join ("
'    query = vbCrLf & query & "select a.custid,promisedate,promisepay from tblnegoptp a where id in"
'    query = vbCrLf & query & "("
'    query = vbCrLf & query & "select id from ( select custid,max(id) as id from tblnegoptp group by 1 ) b"
'    query = vbCrLf & query & "))b on a.custid =b.custid and f_cek_new ilike 'BP%' and agent = 'BP'"
    
    
    query = " select a.*, case when a.promisedate <= b.tgllunas then 1 else 0 end as warna   from ("
    query = vbCrLf & query & " select a.custid, name,promisedate,promisepay, acc_type,tglstatus,agentlama,agent_asli,agent,f_cek_new from mgm a join (select a.custid,promisedate,promisepay from tblnegoptp a where id in(select id from ( select custid,max(id) as id from tblnegoptp group by 1 ) b))b on a.custid =b.custid and f_cek_new ilike 'BP%' and agent = 'BP'"
    If Index <> 0 Then
        query = vbCrLf & query & " and agent_asli = 'TL" & Index & "'"
    End If
    query = vbCrLf & query & " ) a left join vwtbllunaslast b on a.custid = b.custid"
    
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    jmlamt = 0
    
    While Not rs.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = Format(cnull(rs("promisedate")), "YYYY-MM-DD")
             listItem.SubItems(3) = cnull(rs("promisepay"))
             jmlamt = jmlamt + cnull(rs("promisepay"))
             listItem.SubItems(4) = cnull(rs("acc_type"))
             listItem.SubItems(5) = Format(cnull(rs("tglstatus")), "YYYY-MM-DD hh:nn:ss")
             listItem.SubItems(6) = cnull(rs("agentlama"))
             listItem.SubItems(7) = cnull(rs("agent_asli"))
             
             If rs("warna") = 1 Then
                    listItem.ListSubItems(1).ForeColor = vbBlue
                    listItem.ListSubItems(2).ForeColor = vbBlue
                    listItem.ListSubItems(3).ForeColor = vbBlue
                    listItem.ListSubItems(4).ForeColor = vbBlue
                    listItem.ListSubItems(5).ForeColor = vbBlue
                    listItem.ListSubItems(6).ForeColor = vbBlue
                    listItem.ListSubItems(7).ForeColor = vbBlue
                    'listItem.ListSubItems(8).ForeColor = vbBlue
                    'listItem.ListSubItems(9).ForeColor = vbBlue
                End If
             
        rs.MoveNext
    Wend
    Text1.text = rs.RecordCount
    Text2.text = Format(jmlamt, "##,###")
End Sub


Private Sub isilist_log()

    ListView1.ListItems.clear
    
    query = " select * from tbl_log_bp order by tglsend desc limit 300 "
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView1.ListItems.ADD(, , cnull(rs("custid")))
             listItem.SubItems(1) = cnull(rs("name"))
             listItem.SubItems(2) = Format(cnull(rs("ptpdate")), "YYYY-MM-DD")
             listItem.SubItems(3) = cnull(rs("amtptp"))
             listItem.SubItems(4) = cnull(rs("prod"))
             listItem.SubItems(5) = Format(cnull(rs("bpdate")), "YYYY-MM-DD hh:nn:ss")
             listItem.SubItems(6) = cnull(rs("dc"))
             listItem.SubItems(7) = cnull(rs("tl"))
             listItem.SubItems(8) = cnull(rs("sender"))
             listItem.SubItems(9) = Format(cnull(rs("tglsend")), "yyyy-mm-dd")
        rs.MoveNext
    Wend

End Sub

Private Sub Command2_Click()
        
    If Command2.Caption = "Log" Then
        Command2.Caption = "Back"
        headerlist_log
        For i = 0 To 12
            Command1(i).Enabled = False
        Next i
        isilist_log
    Else
        Command2.Caption = "Log"
        headerlist
        For i = 0 To 12
            Command1(i).Enabled = True
        Next i
        isilist
    End If

End Sub

Private Sub Command3_Click()
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data Is Empty!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            cek = cek + 1
        End If
    Next i
        
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If

    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            CustId = ListView1.ListItems(i).text
            Nama = ListView1.ListItems(i).SubItems(1)
            PromiseDate = ListView1.ListItems(i).SubItems(2)
            PromisePay = ListView1.ListItems(i).SubItems(3)
            prod = ListView1.ListItems(i).SubItems(4)
            tglbp = ListView1.ListItems(i).SubItems(5)
            agentlm = ListView1.ListItems(i).SubItems(6)
            agentasl = ListView1.ListItems(i).SubItems(7)
        
            query = "insert into tbl_log_bp values ( '" & CustId & "', '" & Nama & "', '" & PromiseDate & "', '" & PromisePay & "', '" & prod & "', '" & tglbp & "', '" & agentlm & "', '" & agentasl & "', '" & MDIForm1.Text1.text & "' ); " & vbCrLf
            query = query & " update mgm set agent = agent_asli where custid = '" & CustId & "' "
            M_OBJCONN.Execute query
        End If
    Next i
    
    MsgBox "Sended"
End Sub

Private Sub Command4_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If ListView1.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView1.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView1.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView1.ListItems.Count + 1
            For col = 1 To ListView1.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = "'" + ListView1.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = ListView1.ListItems(Row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(Row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_save.ShowOpen
        a = CD_save.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
zzz:
        MsgBox "No data to export", vbInformation, Me.Caption
    End If


End Sub

Private Sub Command5_Click()
    qu = "update mgm set agent_asli = usertbl.team from usertbl where mgm.agentlama = usertbl.userid and mgm.agent = 'BP'"
    M_OBJCONN.Execute qu
    
    Call Form_Load
End Sub

Private Sub Command6_Click()
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data Is Empty!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            cek = cek + 1
        End If
    Next i
        
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If

    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            CustId = ListView1.ListItems(i).text
            Nama = ListView1.ListItems(i).SubItems(1)
            PromiseDate = ListView1.ListItems(i).SubItems(2)
            PromisePay = ListView1.ListItems(i).SubItems(3)
            prod = ListView1.ListItems(i).SubItems(4)
            tglbp = ListView1.ListItems(i).SubItems(5)
            agentlm = ListView1.ListItems(i).SubItems(6)
            agentasl = ListView1.ListItems(i).SubItems(7)
        
            query = "insert into tbl_log_bp values ( '" & CustId & "', '" & Nama & "', '" & PromiseDate & "', '" & PromisePay & "', '" & prod & "', '" & tglbp & "', '" & agentlm & "', '" & agentasl & "', '" & MDIForm1.Text1.text & "' ); " & vbCrLf
            query = query & " update mgm set agent = agentlama where custid = '" & CustId & "' "
            M_OBJCONN.Execute query
        End If
    Next i
    
    MsgBox "Sended"
End Sub

Private Sub Form_Load()
    Call btn
    Call headerlist
    Call isilist
    If MDIForm1.Text2.text = "Supervisor" Or MDIForm1.Text2.text = "Manager" Then
        Command2.Visible = True
        Command3.Visible = True
        Command4.Visible = True
        Command5.Visible = True
        Command6.Visible = True
        Check1.Visible = True
    End If
End Sub

Private Sub btn()
    If MDIForm1.Text2.text <> "Supervisor" And MDIForm1.Text2.text <> "Manager" Then
        
        For i = 0 To 12
            Command1(i).Enabled = False
        Next i
        
        If Left(MDIForm1.Text2.text, 2) = "AM" Then
            query = " select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "'"
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If Len(MDIForm1.Text1.text) = 3 Then
                Command1(Right(MDIForm1.Text1.text, 1)).Enabled = True
            ElseIf Len(MDIForm1.Text1.text) = 4 Then
                Command1(Right(MDIForm1.Text1.text, 2)).Enabled = True
            End If
            For i = 1 To rs.RecordCount
                If Len(rs!tl) = 3 Then
                    a = Right(rs!tl, 1)
                    Command1(a).Enabled = True
                ElseIf Len(rs!tl) = 4 Then
                    a = Right(rs!tl, 2)
                    Command1(a).Enabled = True
                End If
                rs.MoveNext
            Next i
        Else
            query = " select * from  usertbl where userid = '" + MDIForm1.Text1.text + "'"
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If Len(rs!TEAM) = 3 Then
                    a = Right(rs!TEAM, 1)
                    Command1(a).Enabled = True
                ElseIf Len(rs!TEAM) = 4 Then
                    a = Right(rs!TEAM, 2)
                    Command1(a).Enabled = True
                End If
        
        End If
        
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = ListView1.SelectedItem.text
        Me.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub
