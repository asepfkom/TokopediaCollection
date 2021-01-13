VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form formapprovalu_statusptp 
   Caption         =   "Approval Update Status PTP"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10155
   LinkTopic       =   "Form3"
   ScaleHeight     =   5565
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chk_all 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton btntransfer 
      Caption         =   "Transfer"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton btnbatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton btnhstapp 
      Caption         =   "History"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton btnexit 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox cmbapprove 
      Height          =   315
      ItemData        =   "formapprovalu_statusptp.frx":0000
      Left            =   8640
      List            =   "formapprovalu_statusptp.frx":0007
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox cmbsorted 
      Height          =   315
      ItemData        =   "formapprovalu_statusptp.frx":0011
      Left            =   8640
      List            =   "formapprovalu_statusptp.frx":0018
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   8640
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvapprovaltransferdata 
      Height          =   5460
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   9631
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      Caption         =   "Approved by"
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Label2"
      Height          =   255
      Left            =   9000
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Sorted by"
      Height          =   255
      Left            =   8640
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "formapprovalu_statusptp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public objExcel As New Excel.Application
Public objExcelSheet As Excel.Worksheet
Public col, Row, W As Integer
Public a As String

Private Sub HeaderLv()
    lvapprovaltransferdata.ColumnHeaders.ADD , , "No", 600
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Custid", 1100
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Pay Date", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Status PTP Lama", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Status PTP Baru", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "TL Lama", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "TL Baru", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Tanggal Upload", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "PengUpload", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Batch", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "WO Date", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "F CEK NEW", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Pengaprove", 2000
End Sub
Private Sub HeaderLvlv()
    ListView1.ColumnHeaders.ADD , , "No", 600
    ListView1.ColumnHeaders.ADD , , "Custid", 1100
    ListView1.ColumnHeaders.ADD , , "Pay Date", 2000
    ListView1.ColumnHeaders.ADD , , "Status PTP Lama", 2000
    ListView1.ColumnHeaders.ADD , , "Status PTP Baru", 2000
    ListView1.ColumnHeaders.ADD , , "TL Lama", 2000
    ListView1.ColumnHeaders.ADD , , "TL Baru", 2000
    ListView1.ColumnHeaders.ADD , , "Tanggal Upload", 2000
    ListView1.ColumnHeaders.ADD , , "PengUpload", 2000
    ListView1.ColumnHeaders.ADD , , "Batch", 2000
    ListView1.ColumnHeaders.ADD , , "WO Date", 2000
    ListView1.ColumnHeaders.ADD , , "F CEK NEW", 2000
    ListView1.ColumnHeaders.ADD , , "Pengaprove", 2000
    'listview1.ColumnHeaders.ADD , , "Pengaprove", 2000
End Sub

Private Sub HeaderLvhst()
    lvapprovaltransferdata.ColumnHeaders.clear
    lvapprovaltransferdata.ColumnHeaders.ADD , , "No", 600
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Custid", 1100
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Status PTP Lama", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Status PTP Baru", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Tanggal Upload", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Tanggal Transfer", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Pengapprove", 2000
    lvapprovaltransferdata.ColumnHeaders.ADD , , "Pengupload", 2000
End Sub

Private Sub isilv()
    Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
'    sQuery = "select distinct(tbllunas.custid), agentlama, agentbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate from ("
'    sQuery = sQuery + " SELECT a.*, b.recsource, b.B_D, f_cek_new FROM tampungtransferdata a inner join mgm b on a.custid = b.custid) tian "
'    sQuery = sQuery + " left join (select max(paydate) as paydate ,custid from tbllunas group by 2) tbllunas on tian.custid = tbllunas.custid where 1 = 1 and y_n = 1"
    
    'menambahkan TL baru dan TL lama
    
    sQuery = "select custid,statusptplama,statusptpbaru, tllama, z.team tlbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate from ( "
    sQuery = sQuery + " select custid, statusptplama, statusptpbaru, y.team tllama, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate  from "
    sQuery = sQuery + " (select distinct(tian.custid) custid, agentlama, agentbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, "
    sQuery = sQuery + " paydate from ( SELECT a.*, b.recsource, b.B_D, f_cek_new FROM tampungtransferdata a inner join mgm b on a.custid = b.custid) tian  left join "
    sQuery = sQuery + " (select max(paydate) as paydate ,custid from tbllunas group by 2) tbllunas on tian.custid = tbllunas.custid where 1 = 1 and y_n = 1 order by tujapproval) "
    sQuery = sQuery + " x left join usertbl y on x.statusptplama = y.userid) xy left join usertbl z on xy.statusptpbaru = z.userid "
    
    If cmbapprove.text <> "" Then
        sQuery = sQuery + " and tujapproval = '" + cmbapprove.text + "'"
    End If
    sQuery = sQuery + " order by tujapproval"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lvapprovaltransferdata.ListItems.clear
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tanggalupload = Format(RS_Lv("tanggalupload"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = lvapprovaltransferdata.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("paydate")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("statusptplama")))
            listItem.SubItems(4) = Trim(cnull(RS_Lv("statusptpbaru")))
            listItem.SubItems(5) = Trim(cnull(RS_Lv("tllama")))
            listItem.SubItems(6) = Trim(cnull(RS_Lv("tlbaru")))
            listItem.SubItems(7) = tanggalupload
            listItem.SubItems(8) = Trim(cnull(RS_Lv("pengupload")))
            listItem.SubItems(9) = Trim(cnull(RS_Lv("recsource")))
            listItem.SubItems(10) = Trim(cnull(RS_Lv("B_D")))
            listItem.SubItems(11) = Trim(cnull(RS_Lv("f_cek_new")))
            listItem.SubItems(12) = Trim(cnull(RS_Lv("tujapproval")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub Isilvhst()
    Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    sQuery = "SELECT * FROM approvaltransfer order by 5 desc limit 1000"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lvapprovaltransferdata.ListItems.clear
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tanggalupload = Format(RS_Lv("tanggaluploaddaritbltampung"), "yyyy-mm-dd hh:mm:ss")
            tanggaltransfer = Format(RS_Lv("tanggaltransfer"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = lvapprovaltransferdata.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("statusptplama")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("statusptpbaru")))
            listItem.SubItems(4) = tanggalupload
            listItem.SubItems(5) = tanggaltransfer
            listItem.SubItems(6) = Trim(cnull(RS_Lv("pengapprove")))
            listItem.SubItems(7) = Trim(cnull(RS_Lv("penguploaddaritbltampung")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub btnbatal_Click()
    Dim W As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim kosong As Integer
    Dim cmdsql, hst As String
    
    If lvapprovaltransferdata.ListItems.Count = 0 Then
        MsgBox "Data Is Empty!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To lvapprovaltransferdata.ListItems.Count
        If lvapprovaltransferdata.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    a = MsgBox("Yakin Transfer Custid dari Agent Lama ke yang Baru", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Canceled!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For W = 1 To lvapprovaltransferdata.ListItems.Count
        If lvapprovaltransferdata.ListItems(W).Checked = True Then
               
            cmdsql = "DELETE FROM tampungtransferdata WHERE custid ='"
            cmdsql = cmdsql + Trim(lvapprovaltransferdata.ListItems(W).SubItems(1)) + "'"
            M_OBJCONN.execute cmdsql
            
        End If
    Next W
    
    'txt_cust.Text = ""
    Call isilv
End Sub
Private Sub approved()
    Dim sql As String
    Dim M_Objrs As ADODB.Recordset
    
    sql = "select distinct penggaprove from tblpermohonanu_statusptp"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
   cmbapprove.clear
    
    While M_Objrs.EOF = False
        cmbapprove.AddItem CStr(Trim(IIf(IsNull(M_Objrs!penggaprove), "", M_Objrs!penggaprove)))
        M_Objrs.MoveNext
    Wend
End Sub
Private Sub btnexit_Click()
    Unload Me
End Sub

Private Sub btnhstapp_Click()
    If btnhstapp.Caption = "History" Then
        Call HeaderLvhst
        Call Isilvhst
        btnhstapp.Caption = "Back"
        btnhstapp.Top = 240
        btntransfer.Visible = False
        btnbatal.Visible = False
        Label1.Visible = False
        cmbapprove.Visible = False
        Label2.Visible = False
        cmbsorted.Visible = False
    ElseIf btnhstapp.Caption = "Back" Then
        Form_Load
        btnhstapp.Caption = "History"
        btnhstapp.Top = 1200
        btntransfer.Visible = True
        btnbatal.Visible = True
        Label1.Visible = True
        cmbapprove.Visible = True
        Label2.Visible = True
        cmbsorted.Visible = True
    End If
End Sub

Private Sub cetak()
If ListView1.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 
    For col = 1 To ListView1.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = ListView1.ColumnHeaders(col)
    Next
 
    For Row = 2 To ListView1.ListItems.Count + 1
        'If lvapprovaltransferdata.ListItems(Row).Checked = True Then
            For col = 1 To ListView1.ColumnHeaders.Count
                'If lvapprovaltransferdata.ListItems(col).Checked = True Then
                    If col = 1 Then
                            objExcelSheet.Cells(Row, col).Value = ListView1.ListItems(Row - 1).text
                    Else
                        '" 'cararandy 29032016 "
                        Dim hasil1 As String
                            hasil1 = "'" + ListView1.ListItems(Row - 1).SubItems(col - 1)
                            objExcelSheet.Cells(Row, col).Value = hasil1
                    End If
                'End If
            Next
        'End If
    Next
 
    objExcelSheet.Columns.AutoFit
    CD_save.ShowOpen
    a = CD_save.FileName
 
    If a = "" Then
        MsgBox "Export Aborted", vbInformation, Me.Caption
        Exit Sub
    Else
    objExcelSheet.SaveAs a & ".xls"
    MsgBox "Export Completed", vbInformation, Me.Caption
    End If
    objExcel.Workbooks.Open a & ".xls"
    objExcel.Visible = True
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If
End Sub

Private Sub newcetak()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If lvapprovaltransferdata.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
        For col = 1 To lvapprovaltransferdata.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = lvapprovaltransferdata.ColumnHeaders(col)
        Next
     
        For Row = 2 To lvapprovaltransferdata.ListItems.Count + 1
            For col = 1 To lvapprovaltransferdata.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = "'" + lvapprovaltransferdata.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = lvapprovaltransferdata.ListItems(Row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(Row, col).Value = "'" + hasil1
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

Private Sub btntransfer_Click()
    Dim W As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim kosong As Integer
    Dim CMDSQL1, CMDSQL2, cmdsql3, hst As String
    
    If lvapprovaltransferdata.ListItems.Count = 0 Then
        MsgBox "Data Is Empty!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To lvapprovaltransferdata.ListItems.Count
        If lvapprovaltransferdata.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    a = MsgBox("Yakin Transfer Custid dari Agent Lama ke yang Baru", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Canceled!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
        'uploadexceljejaktian========================================================================
'            Call ambilcetak
'            Call cetak
            Call newcetak
        '=============================================================================================
    
    For W = 1 To lvapprovaltransferdata.ListItems.Count
        If lvapprovaltransferdata.ListItems(W).Checked = True Then
            
            'kosong = IIf(IsNull(Format(Trim(lvapprovaltransferdata.ListItems(w).ListSubItems(4)), "yyyy-mm-dd hh:mm:ss")), "", Format(Trim(lvapprovaltransferdata.ListItems(w).ListSubItems(4)), "yyyy-mm-dd hh:mm:ss"))
            'kosong = IIf(IsNull(lvapprovaltransferdata.ListItems(w).ListSubItems(4)), Null, Format(Trim(lvapprovaltransferdata.ListItems(w).ListSubItems(4)), "yyyy-mm-dd hh:mm:ss"))
            'kosong = CStr(IIf(IsNull(M_Objrs!stop_time), "1900-01-01", Format(M_Objrs!stop_time, "yyyy-mm-dd hh:mm:ss")))
            
            CMDSQL1 = "insert into approvalu_statusptp values ('" & lvapprovaltransferdata.ListItems(W).ListSubItems(1) & "'"
            CMDSQL1 = CMDSQL1 + ", '" & lvapprovaltransferdata.ListItems(W).ListSubItems(3) & "', '" & lvapprovaltransferdata.ListItems(W).ListSubItems(4) & "'"
            CMDSQL1 = CMDSQL1 + ", '" & lvapprovaltransferdata.ListItems(W).ListSubItems(7) & "', now(), '" & MDIForm1.Text1.text & "',"
            CMDSQL1 = CMDSQL1 + " '" & lvapprovaltransferdata.ListItems(W).ListSubItems(8) & "')"
            M_OBJCONN.execute CMDSQL1
            
            CMDSQL2 = "update mgm set status_ptp = '" & lvapprovaltransferdata.ListItems(W).ListSubItems(4) & "'"
            CMDSQL2 = CMDSQL2 + " where custid = '" & lvapprovaltransferdata.ListItems(W).ListSubItems(1) & "'"
            'CMDSQL2 = CMDSQL2 + " and agent = '" & lvapprovaltransferdata.ListItems(w).ListSubItems(2) & "'"
            M_OBJCONN.execute CMDSQL2
            
            cmdsql3 = "DELETE FROM tampungtransferdata WHERE custid ='"
            cmdsql3 = cmdsql3 + Trim(lvapprovaltransferdata.ListItems(W).SubItems(1)) + "'"
            M_OBJCONN.execute cmdsql3
        End If
    Next W

    
    'txt_cust.Text = ""
    Call isilv
    ListView1.ListItems.clear
End Sub

Private Sub chk_all_Click()
    Dim r As Integer
        
    If chk_all.Value = vbChecked Then
        If lvapprovaltransferdata.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To lvapprovaltransferdata.ListItems.Count
            lvapprovaltransferdata.ListItems(r).Checked = True
        Next r
        'Call cmd_count_Click
    Else
        For r = 1 To lvapprovaltransferdata.ListItems.Count
            lvapprovaltransferdata.ListItems(r).Checked = False
        Next r
        'Call cmd_count_Click
    End If
End Sub

Private Sub cbaproveklik()
    Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
'    sQuery = "select distinct(custid), agentlama, agentbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval from ("
'    sQuery = sQuery + " SELECT a.*, b.recsource, b.B_D, f_cek_new FROM tampungtransferdata a inner join mgm b on a.custid = b.custid) tian where tujapproval = '" + cmbapprove.Text + "' order by tujapproval"
    
    sQuery = "select custid,statusptplama,statusptpbaru, tllama, z.team tlbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate from ( "
    sQuery = sQuery + " select custid, statusptplama, statusptpbaru, y.team tllama, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate  from "
    sQuery = sQuery + " (select distinct(tian.custid) custid, statusptplama, statusptpbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, "
    sQuery = sQuery + " paydate from ( SELECT a.*, b.recsource, b.B_D, f_cek_new FROM tampungtransferdata a inner join mgm b on a.custid = b.custid) tian  left join "
    sQuery = sQuery + " (select max(paydate) as paydate ,custid from tbllunas group by 2) tbllunas on tian.custid = tbllunas.custid where 1 = 1 and y_n = 1 order by tujapproval) "
    sQuery = sQuery + " x left join usertbl y on x.agentlama = y.userid) xy left join usertbl z on xy.agentbaru = z.userid "
    sQuery = sQuery + " where tujapproval = '" + cmbapprove.text + "' order by tujapproval"
    
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lvapprovaltransferdata.ListItems.clear
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tanggalupload = Format(RS_Lv("tanggalupload"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = lvapprovaltransferdata.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("paydate")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("statusptplama")))
            listItem.SubItems(4) = Trim(cnull(RS_Lv("statusptpbaru")))
            listItem.SubItems(5) = Trim(cnull(RS_Lv("tllama")))
            listItem.SubItems(6) = Trim(cnull(RS_Lv("tlbaru")))
            listItem.SubItems(7) = tanggalupload
            listItem.SubItems(8) = Trim(cnull(RS_Lv("pengupload")))
            listItem.SubItems(9) = Trim(cnull(RS_Lv("recsource")))
            listItem.SubItems(10) = Trim(cnull(RS_Lv("B_D")))
            listItem.SubItems(11) = Trim(cnull(RS_Lv("f_cek_new")))
            listItem.SubItems(12) = Trim(cnull(RS_Lv("tujapproval")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub cmbapprove_Click()
    Call cbaproveklik
End Sub

Private Sub cmbsorted_Click()
    Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    sQuery = "select custid,statusptplama,statusptpbaru, tllama, z.team tlbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate from ( "
    sQuery = sQuery + " select custid, statusptplama, statusptpbaru, y.team tllama, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate  from "
    sQuery = sQuery + " (select distinct(tian.custid) custid, statusptplama, statusptpbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, "
    sQuery = sQuery + " paydate from ( SELECT a.*, b.recsource, b.B_D, f_cek_new FROM tampungupdate_statusptp a inner join mgm b on a.custid = b.custid) tian  left join "
    sQuery = sQuery + " (select max(paydate) as paydate ,custid from tbllunas group by 2) tbllunas on tian.custid = tbllunas.custid where 1 = 1 and y_n = 1 order by tujapproval) "
    sQuery = sQuery + " x left join usertbl y on x.agentlama = y.userid) xy left join usertbl z on xy.agentbaru = z.userid "
    sQuery = sQuery + " where 1 = 1 and y_n = 1 "
    
    If cmbapprove.text <> "" Then
        sQuery = sQuery + " and tujapproval = '" + cmbapprove.text + "'"
    End If
    If cmbsorted.text <> "" Then
        sQuery = sQuery + " and pengupload = '" + cmbsorted.text + "'"
    End If
    sQuery = sQuery + " order by tujapproval"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lvapprovaltransferdata.ListItems.clear
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tanggalupload = Format(RS_Lv("tanggalupload"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = lvapprovaltransferdata.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("paydate")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("statusptplama")))
            listItem.SubItems(4) = Trim(cnull(RS_Lv("statusptpbaru")))
            listItem.SubItems(5) = Trim(cnull(RS_Lv("tllama")))
            listItem.SubItems(6) = Trim(cnull(RS_Lv("tlbaru")))
            listItem.SubItems(7) = tanggalupload
            listItem.SubItems(8) = Trim(cnull(RS_Lv("pengupload")))
            listItem.SubItems(9) = Trim(cnull(RS_Lv("recsource")))
            listItem.SubItems(10) = Trim(cnull(RS_Lv("B_D")))
            listItem.SubItems(11) = Trim(cnull(RS_Lv("f_cek_new")))
            listItem.SubItems(12) = Trim(cnull(RS_Lv("tujapproval")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub cmbsorted_DropDown()
    Dim sql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmbsorted.clear
    
    sql = "select distinct(pengupload) from ("
    sql = sql + " SELECT a.*, b.recsource, b.B_D, f_cek_new FROM tampungupdate_statusptp a inner join mgm b on a.custid = b.custid) tian where 1 = 1 and y_n = 1"
    
    If cmbapprove.text <> "" Then
        sQuery = sQuery + " and tujapproval = '" + cmbapprove.text + "'"
    End If
    sql = sql + " order by pengupload"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
   'cmbapprove.CLEAR
    
    While M_Objrs.EOF = False
        cmbsorted.AddItem CStr(Trim(IIf(IsNull(M_Objrs!pengupload), "", M_Objrs!pengupload)))
        M_Objrs.MoveNext
    Wend
End Sub

Private Sub Form_Load()
    Call HeaderLv
    Call bagitransfer
    Call isilv
    Call HeaderLvlv
    'Call approved
End Sub

Private Sub bagitransfer()
    If (MDIForm1.Text1.text = "JOKO") Or (MDIForm1.Text1.text = "ONTARIO") Or (MDIForm1.Text1.text = "SURJO") Then
        cmbapprove.text = MDIForm1.Text1.text
        cmbapprove.Enabled = False
    End If
End Sub

Private Sub lvapprovaltransferdata_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   lvapprovaltransferdata.SortKey = ColumnHeader.Index - 1
   IndexColumnHEader = ColumnHeader.Index - 1
   lvapprovaltransferdata.Sorted = True
End Sub


Private Sub ambilcetak()
    Dim W As Integer
    Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    
    For W = 1 To lvapprovaltransferdata.ListItems.Count
        If lvapprovaltransferdata.ListItems(W).Checked = True Then
            sQuery = "select custid,statusptplama,statusptpbaru, tllama, z.team tlbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate from ( "
            sQuery = sQuery + " select custid, statusptplama, statusptpbaru, y.team tllama, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, paydate  from "
            sQuery = sQuery + " (select distinct(tian.custid) custid, statusptplama, statusptpbaru, tanggalupload, pengupload, recsource, b_d, f_cek_new, tujapproval, y_n, "
            sQuery = sQuery + " paydate from ( SELECT a.*, b.recsource, b.B_D, f_cek_new FROM tampungupdate_statusptp a inner join mgm b on a.custid = b.custid) tian  left join "
            sQuery = sQuery + " (select max(paydate) as paydate ,custid from tbllunas group by 2) tbllunas on tian.custid = tbllunas.custid where 1 = 1 and y_n = 1 order by tujapproval) "
            sQuery = sQuery + " x left join usertbl y on x.statusptplama = y.userid) xy left join usertbl z on xy.agentbaru = z.userid "
            sQuery = sQuery + " where 1 = 1 and y_n = 1 "
            
            If cmbapprove.text <> "" Then
                sQuery = sQuery + " and tujapproval = '" + cmbapprove.text + "'"
            End If
            If cmbapprove.text <> "" Then
                sQuery = sQuery + " and custid = '" & lvapprovaltransferdata.ListItems(W).ListSubItems(1) & "'"
            End If
            sQuery = sQuery + " order by tujapproval"
            Set RS_Lv = New ADODB.Recordset
            RS_Lv.CursorLocation = adUseClient
            RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
            'ListView1.ListItems.CLEAR
            If RS_Lv.RecordCount > 0 Then
                num = 0
                Do Until RS_Lv.EOF
                    num = num + 1
                    tanggalupload = Format(RS_Lv("tanggalupload"), "yyyy-mm-dd hh:mm:ss")
                    Set listItem = ListView1.ListItems.ADD(, , num)
                    listItem.SubItems(1) = Trim(cnull(RS_Lv("custid")))
                    listItem.SubItems(2) = Trim(cnull(RS_Lv("paydate")))
                    listItem.SubItems(3) = Trim(cnull(RS_Lv("statusptplama")))
                    listItem.SubItems(4) = Trim(cnull(RS_Lv("statusptpbaru")))
                    listItem.SubItems(5) = Trim(cnull(RS_Lv("tllama")))
                    listItem.SubItems(6) = Trim(cnull(RS_Lv("tlbaru")))
                    listItem.SubItems(7) = tanggalupload
                    listItem.SubItems(8) = Trim(cnull(RS_Lv("pengupload")))
                    listItem.SubItems(9) = Trim(cnull(RS_Lv("recsource")))
                    listItem.SubItems(10) = Trim(cnull(RS_Lv("B_D")))
                    listItem.SubItems(11) = Trim(cnull(RS_Lv("f_cek_new")))
                    listItem.SubItems(12) = Trim(cnull(RS_Lv("tujapproval")))
                    RS_Lv.MoveNext
                Loop
            Else
                MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
            End If
        End If
    Next W
    
End Sub

