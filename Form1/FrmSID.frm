VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSID 
   Caption         =   "LIST SID"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   10440
      TabIndex        =   8
      Top             =   6720
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "SEARCH ACCOUNT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command7 
         Caption         =   "Export To Excel (Baru)"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   5880
         Width           =   2655
      End
      Begin VB.TextBox txtlocation 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Import"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   840
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add To List"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export To Excel"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   6360
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxtCustid 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   180
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search by list"
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   2760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label5 
         Caption         =   "Custid:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   6720
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6375
      Left            =   6960
      TabIndex        =   17
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Has Been Export"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5640
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "New Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   7680
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9840
      TabIndex        =   9
      Top             =   6795
      Width           =   855
   End
End
Attribute VB_Name = "FrmSID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs1 As ADODB.Recordset
Private sqlstr As String
Private sCari_list As Boolean

Private Sub Check1_Click()
    Dim xx As Integer
    
    For xx = 1 To ListView1.ListItems.Count
        If Check1.Value = vbChecked Then
            ListView1.ListItems(xx).Checked = True
        Else
            ListView1.ListItems(xx).Checked = False
        End If
    Next xx
End Sub

Private Sub Command1_Click()
    Dim listItem As listItem
    Dim cust_exist As Boolean
    Dim xx As Integer
    Dim list_cust_sel As String
    Dim flag_export As Integer
    
    seq = "select *, date(now()) as hrini from tbl_sequence"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open seq, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        If Format(cnull(rs!Tanggal), "yyyy-mm-dd") <> Format(cnull(rs!hrini), "yyyy-mm-dd") Then
            M_OBJCONN.execute "update tbl_sequence set nomor = '0000', tanggal = now();"
        End If
    End If
    
    sqlstr = "SELECT custid,acc_type,name,dob,ktpno,'" & MDIForm1.Text1.text & "/ID/HBAP/HSBC',mother FROM mgm WHERE custid IS NOT NULL "
    
    If sCari_list = True Then
        For xx = 0 To List1.ListCount - 1
            list_cust_sel = list_cust_sel & "'" & List1.list(xx) & "',"
        Next xx
        
        list_cust_sel = Mid(list_cust_sel, 1, Len(list_cust_sel) - 1)

        sqlstr = sqlstr & " AND custid in (" & list_cust_sel & ")"
    Else
        If TxtCustid.text <> "" Then
            sqlstr = sqlstr & " AND custid like '%" & TxtCustid.text & "%'"
        End If
    End If
    
    M_OBJCONN.execute "DELETE FROM tblcpa_sid;"
    M_OBJCONN.execute "INSERT INTO tblcpa_sid(ref_num,prd_type,name_,dob,id_no,requestor_name,mother_maiden_name) " & sqlstr
    M_OBJCONN.execute "UPDATE tblcpa_sid SET flag_export_sid = '1' WHERE ref_num in (SELECT custid FROM tbl_temp_export_sid)"
    ListView1.ListItems.clear
    
    If rs1.state = 1 Then rs1.Close
    rs1.Open "SELECT * FROM tblcpa_sid;"
    
    If rs1.RecordCount > 0 Then
        cust_exist = False
        If TxtCustid.text <> "" Then
            For xx = 0 To List1.ListCount - 1
                If TxtCustid.text = List1.list(xx) Then
                    cust_exist = True
                End If
            Next xx
            ' Add list customer
            If cust_exist = False Then
                List1.AddItem TxtCustid.text
            End If
        End If
        
        Do Until rs1.EOF
            'ListView1.ForeColor = vbGreen
            Set listItem = ListView1.ListItems.ADD(, , cnull(rs1!Ref_Num))
            listItem.SubItems(1) = cnull(rs1!Prd_type)
            listItem.SubItems(2) = clean_sid(cnull(rs1!name_))
            listItem.SubItems(3) = Format(cnull(rs1!DOB), "DDMMYYYY")
            listItem.SubItems(4) = cnull(rs1!ID_No)
            listItem.SubItems(5) = cnull(rs1!requestor_name)
            listItem.SubItems(6) = cnull(rs1!mother_maiden_name)
            flag_export = cnull(rs1!flag_export_sid)
            If flag_export = 1 Then
                For randy = 1 To 6
                    listItem.ForeColor = vbGreen
                    listItem.ListSubItems(randy).ForeColor = vbGreen
                Next randy
            End If
            rs1.MoveNext
        Loop
    End If
    
    Text1.text = rs1.RecordCount
End Sub

Private Sub new_report_sid()
    xListCustid = ""
    
    For i = 1 To ListView1.ListItems.Count
        xListCustid = xListCustid & "'" & ListView1.ListItems(i).text & "',"
    Next i
    
    xListCustid = Mid(xListCustid, 1, Len(xListCustid) - 1)
    
    vw_sid = "select * from vw_sid where custid in (" & xListCustid & ") "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open vw_sid, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    z = 1
    
    ListView2.ListItems.clear
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Set listItem = ListView2.ListItems.ADD(, , Format(cnull(rs(0)), "dd-mm-yyyy"))
            If Right(Left(cnull(rs(1)), Len(cnull(rs(1))) - 1), 1) = 0 Then
                zz = Right(cnull(rs(1)), 1) + z
                listItem.SubItems(1) = Left(cnull(rs(1)), Len(cnull(rs(1))) - 1) & zz
            Else
                zz = Right(cnull(rs(1)), 2) + z
                listItem.SubItems(1) = Left(cnull(rs(1)), Len(cnull(rs(1))) - 1) & zz
            End If
            listItem.SubItems(2) = cnull(rs(2))
            listItem.SubItems(3) = cnull(rs(3))
            listItem.SubItems(4) = cnull(rs(4))
            listItem.SubItems(5) = cnull(rs(5))
            listItem.SubItems(6) = Format(cnull(rs(6)), "dd-mm-yyyy")
            listItem.SubItems(7) = cnull(rs(7))
            listItem.SubItems(8) = cnull(rs(8))
            listItem.SubItems(9) = cnull(rs(9))
            listItem.SubItems(10) = cnull(rs(10))
            listItem.SubItems(11) = cnull(rs(11))
            listItem.SubItems(12) = cnull(rs(12))
            rs.MoveNext
        Next i
    End If
    
    M_OBJCONN.execute "update tbl_sequence set nomor = '000" & zz & "'"
    
'    If rs1.state = 1 Then rs1.Close
'    rs1.Open "SELECT * FROM (" & _
'             "(SELECT custid as custid_mgm FROM mgm WHERE custid in (" & xListCustid & ")) " & _
'             "AS a LEFT JOIN " & _
'             "(SELECT custid as custid_sid FROM tbl_temp_export_sid WHERE custid in(" & xListCustid & ")) " & _
'             "As b " & _
'             "On a.custid_mgm = b.custid_sid )"
'
'    If rs1.RecordCount > 0 Then
'        Do Until rs1.EOF
'            CustId = cnull(rs1!custid_mgm)
'            custid_sid = cnull(rs1!custid_sid)
'            If custid_sid = "" Then
'                M_OBJCONN.execute "INSERT INTO tbl_temp_export_sid VALUES ('" & CustId & "')"
'            End If
'            rs1.MoveNext
'        Loop
'    End If
End Sub


Private Sub Command2_Click()
    Dim sLokasiExcel As String
    Dim xListCustid As String
    Dim sQuery, CustId, custid_sid As String

    CD_save.Filter = "Excel Files |*.xls"
    CD_save.ShowSave
    
    sLokasiExcel = CD_save.FileName
        
    For i = 1 To ListView1.ListItems.Count
        xListCustid = xListCustid & "'" & ListView1.ListItems(i).text & "',"
    Next i
    
    xListCustid = Mid(xListCustid, 1, Len(xListCustid) - 1)
    
    If rs1.state = 1 Then rs1.Close
    rs1.Open "SELECT * FROM tblcpa_sid WHERE ref_num in (" & xListCustid & ") "
    
    'UPDATE tblcpa_sid
     M_OBJCONN.execute "INSERT INTO tblcpa_sid(ref_num,prd_type,name_,dob,id_no,requestor_name,mother_maiden_name) " & sqlstr

    Call ConvertToExcel_me(rs1, sLokasiExcel)
    
    'RANDY
    If rs1.state = 1 Then rs1.Close
    
    rs1.Open "SELECT * FROM (" & _
             "(SELECT custid as custid_mgm FROM mgm WHERE custid in (" & xListCustid & ")) " & _
             "AS a LEFT JOIN " & _
             "(SELECT custid as custid_sid FROM tbl_temp_export_sid WHERE custid in(" & xListCustid & ")) " & _
             "As b " & _
             "On a.custid_mgm = b.custid_sid )"
    
    If rs1.RecordCount > 0 Then
        Do Until rs1.EOF
            CustId = cnull(rs1!custid_mgm)
            custid_sid = cnull(rs1!custid_sid)
            If custid_sid = "" Then
                M_OBJCONN.execute "INSERT INTO tbl_temp_export_sid VALUES ('" & CustId & "')"
            End If
            rs1.MoveNext
        Loop
    End If
End Sub

Private Sub Command3_Click()
    If List1.ListCount > 0 Then
        sCari_list = True
        Call Command1_Click
        s_cari_list = False
    End If
End Sub

Private Sub Command4_Click()
    Dim cek, K, w, index_hapus As Integer
'    If List1.ListCount > 0 Then
'        List1.RemoveItem List1.ListIndex
'    End If
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "Pilih Data Yang Akan Di-Remove!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
a:
    index_hapus = ListView1.ListItems.Count
    

    For w = 1 To index_hapus
        If ListView1.ListItems(w).Checked = True Then
           ListView1.ListItems.Remove ListView1.ListItems(w).Index
            GoTo a
        End If
    Next w
 
End Sub

Private Sub Command5_Click()
    Dim cst_exist As Boolean

    cst_exist = False
    If TxtCustid.text <> "" Then
        For xx = 0 To List1.ListCount - 1
            If TxtCustid.text = List1.list(xx) Then
                cst_exist = True
            End If
        Next xx
        ' Add list customer
        If cst_exist = False Then
            List1.AddItem TxtCustid.text
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim OBJRECORD As ADODB.Recordset
    Dim ssql As String
    Dim CustId As String
    Dim cmdsql As String
    Dim listItem As listItem
    Dim M_Objrs As ADODB.Recordset
    Dim M_XLSCONN As New ADODB.Connection
    
    On Error GoTo SALAH
    
    With CD_save
        .DialogTitle = "Pilih file excel"
        .Filter = "Excel Files|*.xls"
        .ShowOpen
    End With
        
    txtlocation.text = CD_save.FileName
    
    If CD_save.FileName = "" Then Exit Sub
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CD_save.FileName & ";Extended Properties=Excel 8.0;"
    'M_XLSCONN.OpenSchema (adSchemaTables)
    'Set rs1 = M_XLSCONN.OpenSchema(adSchemaTables)
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    ssql = "SELECT * FROM [Sheet1$] where [custid] is not null"

    OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    If OBJRECORD.RecordCount > 0 Then
        Do Until OBJRECORD.EOF
            List1.AddItem cnull(OBJRECORD!CustId)
            OBJRECORD.MoveNext
        Loop
    End If
    Set OBJRECORD = Nothing
    Exit Sub
SALAH:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Command7_Click()
    Call new_report_sid
    Call export_baru
End Sub

Private Sub export_baru()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    If ListView2.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView2.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView2.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView2.ListItems.Count + 1
            For col = 1 To ListView2.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = ListView2.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = "'" + ListView2.ListItems(Row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(Row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_save.ShowOpen
        a = CD_save.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        'objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
        MsgBox "No data to export", vbInformation, Me.Caption
    End If

End Sub

Private Sub Form_Load()
    Call koneksi
    
    ListView1.ColumnHeaders.ADD , , "Cust ID"
    ListView1.ColumnHeaders.ADD , , "Product"
    ListView1.ColumnHeaders.ADD , , "Name"
    ListView1.ColumnHeaders.ADD , , "Dob"
    ListView1.ColumnHeaders.ADD , , "ID No"
    ListView1.ColumnHeaders.ADD , , "Requestor"
    ListView1.ColumnHeaders.ADD , , "Mother"
    
    ListView2.ColumnHeaders.ADD , , "Capture_date"
    ListView2.ColumnHeaders.ADD , , "Ref_Num"
    ListView2.ColumnHeaders.ADD , , "Informasi Tujuan Permintaan SLIK"
    ListView2.ColumnHeaders.ADD , , "Existing IDHSBC #"
    ListView2.ColumnHeaders.ADD , , "Prd_type"
    ListView2.ColumnHeaders.ADD , , "Name"
    ListView2.ColumnHeaders.ADD , , "DOB"
    ListView2.ColumnHeaders.ADD , , "Place_of_birth"
    ListView2.ColumnHeaders.ADD , , "ID_No"
    ListView2.ColumnHeaders.ADD , , "Requestor Name"
    ListView2.ColumnHeaders.ADD , , "Gender"
    ListView2.ColumnHeaders.ADD , , "Mothers Maiden Name"
    ListView2.ColumnHeaders.ADD , , "Address"
End Sub

Private Sub koneksi()
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenDynamic
    rs1.ActiveConnection = M_OBJCONN
    rs1.LockType = adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs1 = Nothing
End Sub

Private Function clean_sid(sKata As String) As String
    Dim stemp_kata As String
    
    stemp_kata = Trim(sKata)
    stemp_kata = Replace(stemp_kata, "MRS", "")
    stemp_kata = Replace(stemp_kata, "MR", "")
    stemp_kata = Replace(stemp_kata, "MISS", "")
    stemp_kata = Replace(stemp_kata, ".", "")
    stemp_kata = Replace(stemp_kata, ",", "")
    stemp_kata = Trim(stemp_kata)
    
    clean_sid = stemp_kata
End Function

Private Sub ConvertToExcel_me(M_Objrs As ADODB.Recordset, Txtpath As String)
    Dim listItem        As listItem
    Dim cmdsql_update   As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Double
    Dim m_msgbox As String
    
    i = 1
  
    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath = Empty Then
        MsgBox "Nama file tidak boleh kosong, download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    On Error GoTo SALAH
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Double
    If M_Objrs.state = 1 Then
        x = 0
        Y = M_Objrs.fields().Count - 1
        Do Until x > Y
            DoEvents
            objSheet.Cells(1, i).Value = UCase(Replace(CStr(M_Objrs.fields(x).Name), "_", " "))
            i = i + 1
            x = x + 1
        Loop
    End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    'objSheet.Range("A2").CopyFromRecordset M_Objrs '-> Proses pengisian data dimulai dari Cell A2
    
    i = 2
    With M_Objrs
        .MoveFirst
        While Not .EOF
            objSheet.Cells(i, 1).Value = cnull(!Capture_date)
            objSheet.Cells(i, 2).Value = "'" & cnull(!Ref_Num)
            objSheet.Cells(i, 3).Value = "COLECTION_" & Format(cnull(!input_date), "MMDDYYYY")
            objSheet.Cells(i, 4).Value = clean_sid(cnull(!name_))
            objSheet.Cells(i, 5).Value = "'" & Format((cnull(!DOB)), "MM/DD/YYYY")
            objSheet.Cells(i, 6).Value = "'" & cnull(!ID_No)
            objSheet.Cells(i, 7).Value = cnull(!bank1)
            objSheet.Cells(i, 8).Value = cnull(!marketing_source)
            objSheet.Cells(i, 9).Value = cnull(!input_date)
            objSheet.Cells(i, 10).Value = cnull(!requestor_name)
            objSheet.Cells(i, 11).Value = cnull(!user_)
            objSheet.Cells(i, 12).Value = cnull(!result)
            objSheet.Cells(i, 13).Value = cnull(!idi)
            objSheet.Cells(i, 14).Value = cnull(!diff_id)
            objSheet.Cells(i, 15).Value = cnull(!Gender)
            objSheet.Cells(i, 16).Value = cnull(!mother_maiden_name)
            objSheet.Cells(i, 17).Value = cnull(!info_tujuan_permintaan_data_bi)
            i = i + 1
            .MoveNext
        Wend
    End With
    
    objBook.SaveAs Txtpath, xlWorkbookNormal
    objExcel.Quit
    
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    
SALAH:
    Exit Sub
End Sub

