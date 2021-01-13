VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_upstatus_call 
   Caption         =   "Update Status Call"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5520
   LinkTopic       =   "Form3"
   ScaleHeight     =   6750
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Information"
      Height          =   4455
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   5415
      Begin MSComctlLib.ListView LvPTP 
         Height          =   4020
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7091
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "600"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   11160
         Top             =   360
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Check"
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   990
         Width           =   3165
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   3165
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4920
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Caption         =   "Sheet"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Choose File"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_upstatus_call"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M_XLSCONN As New ADODB.Connection

Private Sub cbosheet_Change()
    If txtlocation.text <> "" Then
        If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
        M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.text & "] "
        M_Objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set M_Objrs = Nothing
    End If
End Sub

Private Sub CmdBrowse_Click()
    With CommonDialog1
        .DialogTitle = "Import From File"
        '.Filter = "Excel Files|*.xls;*.xlsx"
        .Filter = "Excel Files|*.xls"
        .ShowOpen
    End With
    txtlocation.text = ""
    If CommonDialog1.FileName = "" Then Exit Sub
    txtlocation.text = CommonDialog1.FileName
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
    Set M_Objrs = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.clear
    If M_Objrs.EOF And M_Objrs.BOF Then Exit Sub
    While Not M_Objrs.EOF
        cbosheet.AddItem IIf(IsNull(M_Objrs!table_name), "", M_Objrs!table_name)
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    'Set M_XLSCONN = Nothing

End Sub

Private Sub Command1_Click()
    qs = "select now() as tanggal"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    data_update = Format(Now, "dd-MM-yyyy")
    abc = Format(M_Objrs!Tanggal, "yyyy_mm_ddi__hhi_mi_ss")
    
    qc = "create table backuptblupdate_sts_call___" & abc & " as (select mgm.custid,mgm.f_cek_new as status_call_lama, a.status_call as status_call_baru from mgm, tblupdate_sts_call a where mgm.custid = a.custid);" & vbCrLf
    qc = qc + "alter table backuptblupdate_sts_call___" & abc & " add column distributeby varchar ;" & vbCrLf
    qc = qc + "alter table backuptblupdate_sts_call___" & abc & " add column tanggal timestamp without time zone default now();" & vbCrLf
    qc = qc + "update backuptblupdate_sts_call___" & abc & " set distributeby = '" & MDIForm1.Text1.text & "';" & vbCrLf & vbCrLf
    qc = qc + "update mgm set f_cek_new = a.status_call from tblupdate_sts_call a where mgm.custid = a.custid;"
    'qc = qc + "insert into tblupdate_sts_call_log(custid,status_call,data_update)select custid,status_call,data_update from tblupdate_sts_call where custid = tblupdate_sts_call.custid;"
    'qc = qc + "update mgm set agent = a.agent from tbltemp_tarik a where mgm.custid = a.custid;" & vbCrLf & vbCrLf
    'qc = qc + "insert into mgm_hst(custid,agent,hst)select custid,agent,hst from tbltemp_tarik where custid = tbltemp_tarik.custid;"
    'qc = qc + "update mgm_hst set agent = a.agent, keterangan = a.hst from tbltemp_tarik a where mgm_hst.custid = a.custid;"
    M_OBJCONN.execute qc
    
    MsgBox "Update Status Call Success"
    
    
End Sub

Private Sub Command2_Click()
    'Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim str_sql2 As String
    Dim RS2 As ADODB.Recordset
    
    qs = "select * from information_schema.columns where table_name = 'tblupdate_sts_call'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
    If M_Objrs.RecordCount = 0 Then
        qc = "create table tblupdate_sts_call ( id serial, custid varchar, status_call varchar );"
        M_OBJCONN.execute qc
    End If
        qd = "delete from tblupdate_sts_call;"
        M_OBJCONN.execute qd
        
        ssql = "SELECT * FROM [" & cbosheet.text & "]   "
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
            
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.CursorType = adOpenDynamic
        rsTemporary.ActiveConnection = M_OBJCONN
        rsTemporary.LockType = adLockOptimistic
            
        rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    
        While Not rs.EOF
        
            str_sql = "INSERT INTO tblupdate_sts_call (custid,status_call) Values ( '" + CStr(rs(0)) + "', '" + rs(1) + "' );"
            M_OBJCONN.execute str_sql
            
            rs.MoveNext
        Wend
                    
                
        a = rs.RecordCount
                
        qs = "select mgm.custid,mgm.name, tblupdate_sts_call.status_call as status_call from mgm,tblupdate_sts_call  where mgm.custid = tblupdate_sts_call.custid"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        LvPTP.ListItems.clear
        
        While Not M_Objrs.EOF
            Set listItem = LvPTP.ListItems.ADD(, , cnull(M_Objrs("custid")))
                listItem.SubItems(1) = cnull(M_Objrs("name"))
'                listItem.SubItems(2) = cnull(M_Objrs("agentlama"))
'                listItem.SubItems(3) = cnull(M_Objrs("agentbaru"))
                listItem.SubItems(2) = cnull(M_Objrs("status_call"))
            M_Objrs.MoveNext
        Wend
        
        b = M_Objrs.RecordCount
        
        
        Dim teks As String
        teks = "Jumlah Data Excel : " & a & vbCrLf & "Jumlah Data Didatabase setelah dicheck : " & b & vbCrLf
        
        If a <> b Then
            teks = teks & "Status : Tidak Sesuai"
        Else
            teks = teks & "Status : Sesuai"
        End If
        MsgBox teks
        Command1.Enabled = True
End Sub

Private Sub Form_Load()
    Call header
End Sub

Private Sub Label3_Click()
    formlogdistribute.Show
End Sub

Private Sub Label5_Click()
    Frame2.Visible = True
End Sub

Private Sub Label7_Click()
    Frame2.Visible = False
End Sub

Private Sub Timer1_Timer()
'    If Label5.BackColor = &H8000000D Then
'        Label5.BackColor = &H8000000F
'    Else
'        Label5.BackColor = &H8000000D
'    End If
End Sub

Private Sub header()
    LvPTP.ColumnHeaders.clear
    With LvPTP.ColumnHeaders
        .ADD 1, , "CUSTID"
        .ADD 2, , "NAMA NASABAH"
'        .ADD 3, , "AGENT LAMA"
'        .ADD 4, , "AGENT DISTRIBUSI"
        .ADD 3, , "STATUS CALL"
    End With
End Sub



