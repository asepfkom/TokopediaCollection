VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRestoreRemarks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete dan Restore Remarks"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13065
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8115
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   14314
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Delete Remark"
      TabPicture(0)   =   "frm_restdelete.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Restore Remark"
      TabPicture(1)   =   "frm_restdelete.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Restore Part Remark"
      TabPicture(2)   =   "frm_restdelete.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Restore Remarks"
         Height          =   6135
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   12495
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   27
            Text            =   "0"
            Top             =   5460
            Width           =   1425
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Proses"
            Height          =   495
            Left            =   10800
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   5460
            Width           =   1275
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   300
            Width           =   5655
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "...."
            Height          =   315
            Left            =   7200
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   300
            Width           =   555
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1470
            TabIndex        =   23
            Text            =   "cbosheet"
            Top             =   660
            Width           =   2355
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   11430
            Top             =   450
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   3645
            Left            =   60
            TabIndex        =   28
            Top             =   1620
            Width           =   12045
            _ExtentX        =   21246
            _ExtentY        =   6429
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label13 
            Caption         =   "List Data Yang Akan di Hapus:"
            Height          =   195
            Left            =   90
            TabIndex        =   33
            Top             =   1200
            Width           =   2205
         End
         Begin VB.Label Label12 
            Caption         =   "Jumlah Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "* File Excel (.xls) Yang Berisi Cust_ID  "
            Height          =   315
            Left            =   7830
            TabIndex        =   31
            Top             =   390
            Width           =   4335
         End
         Begin VB.Label Label10 
            Caption         =   " Source Remarks"
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label Label9 
            Caption         =   "Sheet"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Restore Remark"
         Height          =   7395
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   12435
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   1260
            TabIndex        =   20
            Top             =   4860
            Width           =   1605
         End
         Begin VB.CommandButton cmd_searc 
            Caption         =   "Search..."
            Height          =   375
            Left            =   3120
            TabIndex        =   18
            Top             =   780
            Width           =   975
         End
         Begin VB.CommandButton cmd_rest 
            BackColor       =   &H0080FF80&
            Caption         =   "Restore"
            Height          =   345
            Left            =   3150
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   4830
            Width           =   1005
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frm_restdelete.frx":0054
            Left            =   1140
            List            =   "frm_restdelete.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   810
            Width           =   1905
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   3495
            Left            =   120
            TabIndex        =   15
            Top             =   1230
            Width           =   12045
            _ExtentX        =   21246
            _ExtentY        =   6165
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "Batch Data :"
            Height          =   345
            Left            =   150
            TabIndex        =   21
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Jumlah Data:"
            Height          =   225
            Left            =   180
            TabIndex        =   19
            Top             =   4920
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Delete Remarks"
         Height          =   6135
         Left            =   -74760
         TabIndex        =   2
         Top             =   540
         Width           =   12495
         Begin VB.ComboBox cbosheet 
            Height          =   315
            Left            =   1470
            TabIndex        =   7
            Text            =   "cbosheet"
            Top             =   660
            Width           =   2355
         End
         Begin VB.CommandButton cmdbrowse 
            BackColor       =   &H00C0FFC0&
            Caption         =   "...."
            Height          =   315
            Left            =   7200
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtlocation 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   300
            Width           =   5655
         End
         Begin VB.CommandButton CmdVer 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Proses"
            Height          =   495
            Left            =   10800
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5460
            Width           =   1275
         End
         Begin VB.TextBox txtcount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Text            =   "0"
            Top             =   5460
            Width           =   1425
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   11430
            Top             =   450
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3645
            Left            =   60
            TabIndex        =   8
            Top             =   1620
            Width           =   12045
            _ExtentX        =   21246
            _ExtentY        =   6429
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label4 
            Caption         =   "Sheet"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label3 
            Caption         =   " Source Remarks"
            Height          =   255
            Left            =   60
            TabIndex        =   12
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label Label2 
            Caption         =   "* File Excel (.xls) Yang Berisi Cust_ID  "
            Height          =   315
            Left            =   7830
            TabIndex        =   11
            Top             =   390
            Width           =   4335
         End
         Begin VB.Label Label5 
            Caption         =   "Jumlah Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "List Data Yang Akan di Hapus:"
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   1200
            Width           =   2205
         End
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Delete and Restore Remark"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13035
   End
End
Attribute VB_Name = "FrmRestoreRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_XLSCONN As New ADODB.Connection
Dim sbatch As String
Private Sub cbosheet_Click()
    Dim OBJRECORD As New ADODB.Recordset
    
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
    
        ssql = "SELECT * FROM [" & cboSheet.text & "] "
            rsTemp.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set rsTemp = Nothing
            Set OBJRECORD = New ADODB.Recordset
            OBJRECORD.CursorLocation = adUseClient
            
        ssql = "SELECT * FROM [" & cboSheet.text & "] "
            DoEvents
            OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set DataGrid1.DATASOURCE = OBJRECORD
            txtcount.text = OBJRECORD.RecordCount
End Sub


Private Sub cmd_rest_Click()
    Dim a As String
    
        a = MsgBox("Anda yakin akan Merestore Data Remarks?", vbYesNo + vbInformation, "Konfirmasi")
        If a = vbNo Then
            Exit Sub
        Else
            
'            'balikin mgm
'            str_sql = " update mgm  set f_cek_new = m.f_cek_new tglcall = m.tglcall_hst1, stscallcust = m.stscallcust_hst1, statuscall = m.statuscall_hst1, remarks = m.remarks_hst1, nextactdate = m.nextactdate_hst1, stscallwith = m.statuscallwith_hst1, tglstatus = m.tglstatus_hst1, kethslkerja_new = m.kethslkerja_new_hst1 "
'            str_sql = str_sql + " from deletemgmremake_hst m where custid = id_mgmremarks_hst AND batch_name = '" + Combo1.Text + "' "
'            M_OBJCONN.Execute str_sql

            'balikin mgm
            str_sql = " update mgm  set f_cek_new = m.f_cek_new, tglcall = m.tglcall_hst1, stscallcust = m.stscallcust_hst1, statuscall = m.statuscall_hst1, remarks = m.remarks_hst1, nextactdate = m.nextactdate_hst1, stscallwith = m.statuscallwith_hst1, tglstatus = m.tglstatus_hst1, kethslkerja_new = m.kethslkerja_new_hst1 "
            'addby tian req doddy 14Feb2017
            str_sql = str_sql + " ,HOMENOADD1 = m.HOMENOADD1, homenoadd2 = m.homenoadd2, officenoadd1 = m.officenoadd1, officenoadd2 = m.officenoadd2, mobilenoadd1 = m.mobilenoadd1, MOBILENOADD2 = m.MOBILENOADD2,stskathomeadd1 = m.stskathomeadd1,stskathomeadd2 = m.stskathomeadd2,stskatofficeadd1 = m.stskatofficeadd1,stskatofficeadd2 = m.stskatofficeadd2,stskathpadd1 = m.stskathpadd1 ,stskathpadd2 = m.stskathpadd2 "
            str_sql = str_sql + " from deletemgmremake_hst m where custid = id_mgmremarks_hst AND batch_name = '" + Combo1.text + "' "
            M_OBJCONN.Execute str_sql
                
            'balikin mgm_hst
            str_sql = "insert into mgm_hst (f_cek_new, custid, tgl,agent, hst, kodeds, kdcomplaint, f_cek, statuscall, ststelpwith, user_log) select "
            str_sql = str_sql + "f_cek_new, id_restoremgmhst, restore_tgl, restore_agent, restore_hst, restore_kodeds, restore_kdcomplaint,"
            str_sql = str_sql + "restore_f_cek, restore_statuscall, restore_ststelpwith, restore_user_log from restoremgmhst where id_restoremgmhst in ( select custid from tbl_uploadexcel where batch_name = '" + Combo1.text + "')"
            M_OBJCONN.Execute str_sql
            
            'restore lognya
            M_OBJCONN.Execute "insert into restoremgmhst_log select * from restoremgmhst where id_restoremgmhst in ( select custid from tbl_uploadexcel where batch_name = '" + Combo1.text + "')"
            
            M_OBJCONN.Execute "delete from restoremgmhst where restore_batch = '" & Combo1.text & "'"
            
            
        End If
        
        MsgBox "Data berhasil Di Restore !"
        Unload Me
End Sub

Private Sub cmd_searc_Click()
    If Combo1.text = "" Then
        MsgBox "Masukkan Batch Number Terlebih Dahulu !"
        Exit Sub
    Else
        str_sql = "select * from restoremgmhst where restore_batch = '" + Combo1.text + "' "
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open str_sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
           
        Set DataGrid2.DATASOURCE = rs
        Text1.text = rs.RecordCount
        
        cmd_rest.Enabled = True
    End If
End Sub

Private Sub CmdVer_Click()
    Dim rs As New ADODB.Recordset
    Dim temp_rs As ADODB.Recordset
    Dim str_sql As String
    Dim scustid As String
    
        If CommonDialog1.FileName = "" Then
            MsgBox "Browse Data Excel Terlebih Dahulu", vbInformation + vbOKOnly, "Information"
            Exit Sub
        End If
        
        If cboSheet.text = "" Then
           MsgBox "Pilih Sheet", vbInformation + vbOKOnly, "Information"
           cboSheet.SetFocus
           Exit Sub
        End If

'        str_sql = "select * from deletemgmremake_hst where batch_name = '" + sbatch + "'"
'        Set RS = New ADODB.Recordset
'        RS.CursorLocation = adUseClient
'        RS.Open str_sql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        If RS.RecordCount > 0 Then
'            MsgBox "Anda Telah Pernah Melakukan Proses!"
'
'            Exit Sub
'        End If
        
'        Set RS = Nothing
        
        ssql = "SELECT * FROM [" & cboSheet.text & "]   "
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.CursorType = adOpenDynamic
        rsTemporary.ActiveConnection = M_OBJCONN
        rsTemporary.LockType = adLockOptimistic
        
        rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        M_OBJCONN.Execute "delete from tbl_uploadexcel"
        While Not rs.EOF
            scustid = IIf(IsNull(rs("custid")), "", rs("custid"))
                str_sql = "INSERT INTO tbl_uploadexcel(custid, batch_name) values ('" + scustid + "', '" + sbatch + "')   "
                M_OBJCONN.Execute str_sql
            rs.MoveNext
        Wend
        
        Set rsTemporary = Nothing
        
'        'Backup mgm_hst
'        str_sql = "insert into restoremgmhst (f_cek_new, id_restoremgmhst, restore_tgl, restore_agent, restore_hst, restore_kodeds, restore_kdcomplaint, "
'        str_sql = str_sql & " restore_f_cek, restore_statuscall, restore_ststelpwith, restore_user_log, restore_batch) select"
'        str_sql = str_sql & " f_cek_new, custid, tgl,agent, hst, kodeds, kdcomplaint, f_cek, statuscall, ststelpwith,"
'        str_sql = str_sql & " user_log, '" & sbatch & "' from mgm_hst where custid in (select custid from tbl_uploadexcel)"
'        M_OBJCONN.Execute str_sql
'
'        'Backup mgm
'        str_sql = "insert into deletemgmremake_hst(f_cek_new,id_mgmremarks_hst, tglcall_hst1, stscallcust_hst1, statuscall_hst1, remarks_hst1, batch_name, nextactdate_hst1, statuscallwith_hst1, tglstatus_hst1, kethslkerja_new_hst1) "
'        str_sql = str_sql + " select f_cek_new, custid, tglcall, stscallcust, statuscall,remarks,'" & sbatch & "' ,nextactdate,stscallwith,tglstatus,kethslkerja_new  "
'        str_sql = str_sql + " from mgm where custid in (select custid from tbl_uploadexcel) AND custid not in (select custid from deletemgmremake_hst)"
'        M_OBJCONN.Execute str_sql
'
'        'Clear mgm
'        str_sql = "update mgm set f_cek_new = null, statuscall = null, stscallwith = null, tglcall = null, nextactdate = null, tglstatus = null, stscallcust = null, kethslkerja_new = null, remarks = null"
'        str_sql = str_sql + " where custid in (select custid from tbl_uploadexcel)"
'        M_OBJCONN.Execute str_sql
'
'        'Clear mgm_hst
'        str_sql = "delete from mgm_hst where custid in (select custid from tbl_uploadexcel)"
'        M_OBJCONN.Execute str_sql

        'Backup mgm_hst
        str_sql = "insert into restoremgmhst (f_cek_new, id_restoremgmhst, restore_tgl, restore_agent, restore_hst, restore_kodeds, restore_kdcomplaint, "
        str_sql = str_sql & " restore_f_cek, restore_statuscall, restore_ststelpwith, restore_user_log, restore_batch) select"
        str_sql = str_sql & " f_cek_new, custid, tgl,agent, hst||phoneno, kodeds, kdcomplaint, f_cek, statuscall, ststelpwith,"
        str_sql = str_sql & " user_log, '" & sbatch & "' from mgm_hst where custid in (select custid from tbl_uploadexcel)"
        M_OBJCONN.Execute str_sql
        
        'Backup mgm
        str_sql = "insert into deletemgmremake_hst(f_cek_new,id_mgmremarks_hst, tglcall_hst1, stscallcust_hst1, statuscall_hst1, remarks_hst1, batch_name, nextactdate_hst1, statuscallwith_hst1, tglstatus_hst1, kethslkerja_new_hst1"
        'addby tian req doddy 14Feb2017
        str_sql = str_sql + " ,HOMENOADD1,homenoadd2,officenoadd1,officenoadd2,mobilenoadd1,MOBILENOADD2,stskathomeadd1,stskathomeadd2,stskatofficeadd1,stskatofficeadd2,stskathpadd1,stskathpadd2 ) "
        str_sql = str_sql + " select f_cek_new, custid, tglcall, stscallcust, statuscall,remarks,'" & sbatch & "' ,nextactdate,stscallwith,tglstatus,kethslkerja_new  "
        'addby tian req doddy 14Feb2017
        str_sql = str_sql + " ,HOMENOADD1,homenoadd2,officenoadd1,officenoadd2,mobilenoadd1,MOBILENOADD2,stskathomeadd1,stskathomeadd2,stskatofficeadd1,stskatofficeadd2,stskathpadd1,stskathpadd2 "
        str_sql = str_sql + " from mgm where custid in (select custid from tbl_uploadexcel) AND custid not in (select id_mgmremarks_hst from deletemgmremake_hst)"
        M_OBJCONN.Execute str_sql
            
        'Clear mgm
        str_sql = "update mgm set f_cek_new = null, statuscall = null, stscallwith = null, tglcall = null, nextactdate = null, tglstatus = null, stscallcust = null, kethslkerja_new = null, remarks = null"
        'addby tian req doddy 14Feb2017
        str_sql = str_sql + " ,HOMENOADD1 = null,homenoadd2 = null,officenoadd1 = null,officenoadd2 = null,mobilenoadd1 = null,MOBILENOADD2 = null, stskathomeadd1 = null, stskathomeadd2 = null, stskatofficeadd1 = null, stskatofficeadd2 = null, stskathpadd1 = null, stskathpadd2 = null, spv_allow = now() "
        str_sql = str_sql + " where custid in (select custid from tbl_uploadexcel)"
        M_OBJCONN.Execute str_sql
        
        'Clear mgm_hst
        str_sql = "delete from mgm_hst where custid in (select custid from tbl_uploadexcel)"
        M_OBJCONN.Execute str_sql
                
        ' UPDATE mgm SET Tgl null etc
       
        'Set M_XLSCONN = Nothing
        Set rs = Nothing
        MsgBox "Data Remark Berhasil Dihapus!"
        Unload Me
        'cmddel.Enabled = True
        
        
End Sub


Private Sub CmdBrowse_Click()
    With CommonDialog1
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
    End With
        
    txtlocation.text = CommonDialog1.FileName
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtlocation.text & ";Extended Properties=Excel 8.0;"
    Set rsTemp = M_XLSCONN.OpenSchema(adSchemaTables)
    cboSheet.clear
    If rsTemp.EOF And rsTemp.BOF Then Exit Sub
        
    While Not rsTemp.EOF
        cboSheet.AddItem IIf(IsNull(rsTemp!table_name), "", rsTemp!table_name)
        rsTemp.MoveNext
    Wend
    
    Set rsTemp = Nothing
End Sub


Private Sub cmddel_Click()
'    Dim a As String
'
'        a = MsgBox("Anda yakin akan Menghapus Data Remarks?", vbYesNo + vbInformation, "Konfirmasi")
'        If a = vbNo Then
'            Exit Sub
'        End If
'
'        If txtlocation.Text = "" Then
'            MsgBox "Source custid masih kosong!", vbOKOnly + vbInformation, "Informasi"
'            Exit Sub
'        End If
'
'        If cbosheet.Text = "" Then
'            MsgBox "Sheet masih kosong!", vbOKOnly + vbInformation, "Informasi"
'            Exit Sub
'        End If
'
'
'
'
'        Unload Me
End Sub


Private Sub Combo2_Click()
        Dim OBJRECORD As New ADODB.Recordset
    
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
    
        ssql = "SELECT * FROM [" & Combo2.text & "] "
            rsTemp.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set rsTemp = Nothing
            Set OBJRECORD = New ADODB.Recordset
            OBJRECORD.CursorLocation = adUseClient
            
        ssql = "SELECT * FROM [" & Combo2.text & "] "
            DoEvents
            OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set DataGrid3.DATASOURCE = OBJRECORD
            Text3.text = OBJRECORD.RecordCount
End Sub

Private Sub Command1_Click()
    With CommonDialog2
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
    End With
        
    Text2.text = CommonDialog2.FileName
    
    If CommonDialog2.FileName = "" Then Exit Sub
    
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & Text2.text & ";Extended Properties=Excel 8.0;"
    Set rsTemp = M_XLSCONN.OpenSchema(adSchemaTables)
    Combo2.clear
    If rsTemp.EOF And rsTemp.BOF Then Exit Sub
        
    While Not rsTemp.EOF
        Combo2.AddItem IIf(IsNull(rsTemp!table_name), "", rsTemp!table_name)
        rsTemp.MoveNext
    Wend
    
    Set rsTemp = Nothing
End Sub

Private Sub Command2_Click()
    Dim a As String

        If CommonDialog2.FileName = "" Then
            MsgBox "Browse Data Excel Terlebih Dahulu", vbInformation + vbOKOnly, "Information"
            Exit Sub
        End If
        
        If Combo2.text = "" Then
           MsgBox "Pilih Sheet", vbInformation + vbOKOnly, "Information"
           cboSheet.SetFocus
           Exit Sub
        End If
    
    
    
        a = MsgBox("Anda yakin akan Merestore Data Remarks?", vbYesNo + vbInformation, "Konfirmasi")
        If a = vbNo Then
            Exit Sub
        Else
            
        ssql = "SELECT * FROM [" & Combo2.text & "]   "
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.CursorType = adOpenDynamic
        rsTemporary.ActiveConnection = M_OBJCONN
        rsTemporary.LockType = adLockOptimistic
        
        rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        M_OBJCONN.Execute "delete from tbl_uploadexcel_restore"
        While Not rs.EOF
            scustid = IIf(IsNull(rs("custid")), "", rs("custid"))
                str_sql = "INSERT INTO tbl_uploadexcel_restore(custid) values ('" + scustid + "')   "
                M_OBJCONN.Execute str_sql
            rs.MoveNext
        Wend
        
        Set rsTemporary = Nothing

            'balikin mgm
            str_sql = " update mgm  set f_cek_new = m.f_cek_new, tglcall = m.tglcall_hst1, stscallcust = m.stscallcust_hst1, statuscall = m.statuscall_hst1, remarks = m.remarks_hst1, nextactdate = m.nextactdate_hst1, stscallwith = m.statuscallwith_hst1, tglstatus = m.tglstatus_hst1, kethslkerja_new = m.kethslkerja_new_hst1 "
            'addby tian req doddy 14Feb2017
            str_sql = str_sql + " ,HOMENOADD1 = m.HOMENOADD1, homenoadd2 = m.homenoadd2, officenoadd1 = m.officenoadd1, officenoadd2 = m.officenoadd2, mobilenoadd1 = m.mobilenoadd1, MOBILENOADD2 = m.MOBILENOADD2,stskathomeadd1 = m.stskathomeadd1,stskathomeadd2 = m.stskathomeadd2,stskatofficeadd1 = m.stskatofficeadd1,stskatofficeadd2 = m.stskatofficeadd2,stskathpadd1 = m.stskathpadd1 ,stskathpadd2 = m.stskathpadd2 "
            str_sql = str_sql + " from (select * from deletemgmremake_hst where id_mgmremarks_hst in (select custid from tbl_uploadexcel_restore)) m where custid = id_mgmremarks_hst "
            M_OBJCONN.Execute str_sql
                
            'balikin mgm_hst
            str_sql = "insert into mgm_hst (f_cek_new, custid, tgl,agent, hst, kodeds, kdcomplaint, f_cek, statuscall, ststelpwith, user_log) select "
            str_sql = str_sql + "f_cek_new, id_restoremgmhst, restore_tgl, restore_agent, restore_hst, restore_kodeds, restore_kdcomplaint,"
            str_sql = str_sql + "restore_f_cek, restore_statuscall, restore_ststelpwith, restore_user_log from restoremgmhst where id_restoremgmhst in (select custid from tbl_uploadexcel_restore)"
            M_OBJCONN.Execute str_sql
            
            'restore lognya
            M_OBJCONN.Execute "insert into restoremgmhst_log select * from restoremgmhst where id_restoremgmhst in (select custid from tbl_uploadexcel_restore) "
            
            M_OBJCONN.Execute "delete from restoremgmhst where id_restoremgmhst in (select custid from tbl_uploadexcel_restore)"
            
            
        End If
        
        MsgBox "Data berhasil Di Restore !"
        Unload Me
        
End Sub

Private Sub Form_Load()
    cmd_rest.Enabled = False
    
    str_sql = "select distinct restore_batch from restoremgmhst order by restore_batch"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open str_sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
    While Not rs.EOF
        Combo1.AddItem IIf(IsNull(rs!restore_batch), "", rs!restore_batch)
        rs.MoveNext
    Wend
    
    sbatch = Format(Now, "ddmmyyyy")
End Sub


