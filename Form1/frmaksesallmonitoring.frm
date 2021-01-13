VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmaksesallmonitoring 
   Caption         =   "Monitoring AccessALL"
   ClientHeight    =   7560
   ClientLeft      =   420
   ClientTop       =   765
   ClientWidth     =   10215
   LinkTopic       =   "Form5"
   ScaleHeight     =   7560
   ScaleWidth      =   10215
   Begin VB.CommandButton Command2 
      Caption         =   "Export"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin MSComctlLib.ListView LvAcc 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   12303
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
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7200
      Width           =   1935
   End
End
Attribute VB_Name = "frmaksesallmonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HeaderAgent()
    LvAcc.ColumnHeaders.clear
    LvAcc.ColumnHeaders.ADD 1, , "BATCH", 2000
    LvAcc.ColumnHeaders.ADD 2, , "CUSTID", 2000
    LvAcc.ColumnHeaders.ADD 3, , "NAMA", 2000
    LvAcc.ColumnHeaders.ADD 4, , "AKSES OLEH", 2000
    LvAcc.ColumnHeaders.ADD 5, , "TOUCH", 2000
End Sub

Private Sub IsiAccount()
    
'    cmdsql = "select kd_profile, a.custid, a.name, monitor_akses, touch  from mgm a, " & vbCrLf
'    cmdsql = cmdsql + " (" & vbCrLf
'    cmdsql = cmdsql + " select *,count(user_log) as touch  from (" & vbCrLf
'    cmdsql = cmdsql + " select distinct kd_profile, custid,user_log from (" & vbCrLf
'    cmdsql = cmdsql + " select a.*, b.kd_profile from mgm_hst a," & vbCrLf
'    cmdsql = cmdsql + " (select a.*,b.custid from tbl_profile_aksesall a, tbl_cust_aksesall b where a.kd_profile =b.kd_profile) b" & vbCrLf
'    cmdsql = cmdsql + " where a.custid = b.custid and tgl between b.waktu_awal and b.waktu_akhir" & vbCrLf
'    cmdsql = cmdsql + " ) c ) c group by 1,2,3" & vbCrLf
'    cmdsql = cmdsql + " ) b" & vbCrLf
'    cmdsql = cmdsql + " where agent = 'AKSESALL' and a.custid = b.custid"

    cmdsql = " select a.kd_profile, a.custid, a.name, monitor_akses, touch, agent from (" & vbCrLf
    cmdsql = cmdsql + " select b.kd_profile,a.custid,a.name,a.monitor_akses,a.agent  from mgm a,(select a.kd_profile,b.custid from tbl_profile_aksesall a, tbl_cust_aksesall b where a.kd_profile =b.kd_profile" & vbCrLf
    cmdsql = cmdsql + " ) b where a.custid = b.custid" & vbCrLf
    cmdsql = cmdsql + " ) a left join" & vbCrLf
    cmdsql = cmdsql + " (" & vbCrLf
    cmdsql = cmdsql + " select kd_profile,custid,count(user_log) as touch  from (" & vbCrLf
    cmdsql = cmdsql + " select distinct kd_profile, custid,user_log from (" & vbCrLf
    cmdsql = cmdsql + " select a.*, b.kd_profile from mgm_hst a," & vbCrLf
    cmdsql = cmdsql + " (select a.*,b.custid from tbl_profile_aksesall a, tbl_cust_aksesall b where a.kd_profile =b.kd_profile) b" & vbCrLf
    cmdsql = cmdsql + " where a.custid = b.custid and tgl between b.waktu_awal and b.waktu_akhir" & vbCrLf
    cmdsql = cmdsql + ") c ) c group by 1,2" & vbCrLf
    cmdsql = cmdsql + " ) b" & vbCrLf
    cmdsql = cmdsql + " on a.custid = b.custid and a.kd_profile = b.kd_profile where agent = 'AKSESALL'"

    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.clear
    Label1.Caption = "Jumlah Data : " & M_Objrs.RecordCount
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    Else
        For i = 1 To M_Objrs.RecordCount
            Set listItem = LvAcc.ListItems.ADD(, , M_Objrs("kd_profile"))
            listItem.SubItems(1) = M_Objrs("custid")
            listItem.SubItems(2) = IIf(IsNull(M_Objrs("name")), "", M_Objrs("name"))
            listItem.SubItems(3) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            listItem.SubItems(4) = IIf(IsNull(M_Objrs("touch")), "", M_Objrs("touch"))
            M_Objrs.MoveNext
        Next i
    End If
    Set M_Objrs = Nothing
End Sub

Private Sub Command1_Click()
    Call Form_Load
End Sub

Private Sub Command2_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If LvAcc.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To LvAcc.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = LvAcc.ColumnHeaders(col)
        Next
     
        For Row = 2 To LvAcc.ListItems.Count + 1
            For col = 1 To LvAcc.ColumnHeaders.Count
                If col = 1 Then
                        objExcelSheet.Cells(Row, col).Value = "'" + LvAcc.ListItems(Row - 1).text
                ElseIf col = 2 Then
                        objExcelSheet.Cells(Row, col).Value = "'" + LvAcc.ListItems(Row - 1).text
                Else
                    '" 'cararandy 29032016 "
                    Dim hasil1 As String
                        hasil1 = LvAcc.ListItems(Row - 1).SubItems(col - 1)
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

Private Sub Form_Load()
    Call HeaderAgent
    Call IsiAccount
End Sub

Private Sub LvAcc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAcc.SortKey = ColumnHeader.Index - 1
    LvAcc.Sorted = True
End Sub

Private Sub LvAcc_DblClick()
    If LvAcc.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).text = LvAcc.SelectedItem.SubItems(1)
        Me.Hide
        VIEW_MGMDATA.Show
    
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub
