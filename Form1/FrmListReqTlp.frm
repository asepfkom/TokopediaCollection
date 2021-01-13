VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmListReqTlp 
   Caption         =   "List Request Number Telephone"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Request Number"
      TabPicture(0)   =   "FrmListReqTlp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LstReq"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdApprove"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdCekAll"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdUnCekAll"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtReq"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdReject"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Log Request Number"
      TabPicture(1)   =   "FrmListReqTlp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LstReqLog"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtReqLog"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command1 
         Caption         =   "List &Reject"
         Enabled         =   0   'False
         Height          =   555
         Left            =   7080
         TabIndex        =   11
         Top             =   5340
         Width           =   1575
      End
      Begin VB.CommandButton CmdReject 
         Caption         =   "&Reject"
         Height          =   555
         Left            =   4980
         TabIndex        =   10
         Top             =   5340
         Width           =   1575
      End
      Begin VB.TextBox TxtReqLog 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -65160
         TabIndex        =   9
         Text            =   "0"
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox TxtReq 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9840
         TabIndex        =   7
         Text            =   "0"
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton CmdUnCekAll 
         Caption         =   "&UnCek All"
         Height          =   555
         Left            =   1740
         TabIndex        =   5
         Top             =   5340
         Width           =   1575
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "&Cek All"
         Height          =   555
         Left            =   180
         TabIndex        =   4
         Top             =   5340
         Width           =   1575
      End
      Begin VB.CommandButton CmdApprove 
         Caption         =   "&Approve"
         Height          =   555
         Left            =   3420
         TabIndex        =   3
         Top             =   5340
         Width           =   1575
      End
      Begin MSComctlLib.ListView LstReq 
         Height          =   4860
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   8573
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
      Begin MSComctlLib.ListView LstReqLog 
         Height          =   4860
         Left            =   -74820
         TabIndex        =   2
         Top             =   480
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   8573
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
      Begin VB.Label Label2 
         Caption         =   "Jumlah data:"
         Height          =   255
         Left            =   -66180
         TabIndex        =   8
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah data:"
         Height          =   255
         Left            =   8820
         TabIndex        =   6
         Top             =   5460
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmListReqTlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HeaderReq()
    LstReq.ColumnHeaders.ADD 1, , "ID", 800
    LstReq.ColumnHeaders.ADD 2, , "Tgl.Request", 0
    LstReq.ColumnHeaders.ADD 3, , "Custid", 2000
    LstReq.ColumnHeaders.ADD 4, , "Agent", 1500
    LstReq.ColumnHeaders.ADD 5, , "Home 1", 0
    LstReq.ColumnHeaders.ADD 6, , "Home 2", 0
    LstReq.ColumnHeaders.ADD 7, , "Office 1", 0
    LstReq.ColumnHeaders.ADD 8, , "Office 2", 0
    LstReq.ColumnHeaders.ADD 9, , "Mobile 1", 0
    LstReq.ColumnHeaders.ADD 10, , "Mobile 2", 0
    LstReq.ColumnHeaders.ADD 11, , "EcPhone", 0
    
    '@@17042012, Perubahan u/ request number hanya ada nomor dan kategori
    LstReq.ColumnHeaders.ADD 12, , "Request Number", 1500
    LstReq.ColumnHeaders.ADD 13, , "Kategori", 3000
    
    LstReq.ColumnHeaders.ADD 14, , "Keterangan", 3000
    LstReq.ColumnHeaders.ADD 15, , "Jenis", 3000
End Sub

Private Sub HeaderLstReq()
    LstReqLog.ColumnHeaders.ADD 1, , "Tgl.Request", 1500
    LstReqLog.ColumnHeaders.ADD 2, , "Custid", 1500
    LstReqLog.ColumnHeaders.ADD 3, , "Agent", 1500
    LstReqLog.ColumnHeaders.ADD 4, , "Home 1", 1500
    LstReqLog.ColumnHeaders.ADD 5, , "Home 2", 1500
    LstReqLog.ColumnHeaders.ADD 6, , "Office 1", 1500
    LstReqLog.ColumnHeaders.ADD 7, , "Office 2", 1500
    LstReqLog.ColumnHeaders.ADD 8, , "Mobile 1", 1500
    LstReqLog.ColumnHeaders.ADD 9, , "Mobile 2", 1500
    LstReqLog.ColumnHeaders.ADD 10, , "EcPhone", 1500
    LstReqLog.ColumnHeaders.ADD 11, , "Tgl.Approve", 1500
    LstReqLog.ColumnHeaders.ADD 12, , "Approve By", 1500
    LstReqLog.ColumnHeaders.ADD 13, , "Keterangan", 3000
End Sub

Private Sub CmdApprove_Click()
    Dim M_DATA      As New CLS_FRMCUST_CC_MGM
    Dim CMDSQL      As String
    Dim M_objrs     As ADODB.Recordset
    Dim W           As Integer
    Dim ListItem    As ListItem
    Dim pesan       As String
    Dim K           As String
    Dim strket_hst  As String
    Dim bAdd_phone  As Boolean
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    K = MsgBox("Anda yakin akan melakukan approve number?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If K = vbNo Then
        Exit Sub
    End If
    
    cmdapprove.Enabled = False
    CmdCekAll.Enabled = False
    CmdUncekAll.Enabled = False
    CmdReject.Enabled = False
    
    DoEvents
    bAdd_phone = False
    For W = 1 To LstReq.ListItems.Count
        If LstReq.ListItems(W).Checked = True Then
            
'            '// CEK NO TELP EXISTS ///////////////////////////////////////
'            If M_Objrs.state = 1 Then M_Objrs.Close
'            M_Objrs.Open "SELECT custid FROM mgm_hst WHERE date_part('year',tgl)=date_part('year',now()) AND phoneno='" & CStr(Trim(LstReq.ListItems(W).SubItems(11))) & "' LIMIT 1 "
            
            '// REQUEST CANCEL
'            If M_Objrs.RecordCount > 0 Then
'                MsgBox "No Telp " & CStr(LstReq.ListItems(W).SubItems(11)) & " Pada CH " & M_Objrs("custid") & " Permintaan no telp dibatalkan !", vbOKOnly + vbInformation, "INFO"
                
                'Update Data Ke tabel LOg Telepon
'                cmdsql = "insert into tblrequestadditionalphone_log "
'                cmdsql = cmdsql + "select * from tblrequestadditionalphone where id='"
'                cmdsql = cmdsql + LstReq.ListItems(W).Text + "'"
'                M_OBJCONN.Execute cmdsql
'
'                'Update data log, tgl approve dan di approve oleh
'                cmdsql = "update tblrequestadditionalphone_log set tglapprove='"
'                cmdsql = cmdsql + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss") + "', approve_by='"
'                cmdsql = cmdsql + MDIForm1.Text1.Text + "',keterangan=keterangan||' ::Request dibatalkan:::' where id='"
'                cmdsql = cmdsql + LstReq.ListItems(W).Text + "'"
'                M_OBJCONN.Execute cmdsql
'
'                'Hapus data di tabel tblrequestadditionalphone
'                cmdsql = "delete from tblrequestadditionalphone where id='"
'                cmdsql = cmdsql + LstReq.ListItems(W).Text + "'"
'                M_OBJCONN.Execute cmdsql
            
'            Else
                ' Jika duplikat di lompat 06 Agustus 2014
'                On Error GoTo next_step
'                M_OBJCONN.Execute ("INSERT INTO tbllist_addphone (custid,phone) VALUES ('" & LstReq.ListItems(w).SubItems(2) & "','" & CStr(Trim(LstReq.ListItems(w).SubItems(11))) & "')")
                
                'If bAdd_phone Then
                pesan = "Request Number di Approve: " & vbCrLf
                pesan = pesan & " Custid : " & LstReq.ListItems(W).SubItems(2) & vbCrLf
                pesan = pesan & " Di approve oleh : " & MDIForm1.Text1.text
                
                'Update di MGM
                If LstReq.ListItems(W).SubItems(14) = "AddHome1" Then
                    CMDSQL = "update mgm set homenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathomeadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddHome2" Then
                    CMDSQL = "update mgm set homenoadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathomeadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddOffice1" Then
                    CMDSQL = "update mgm set officenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskatofficeadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                 ElseIf LstReq.ListItems(W).SubItems(14) = "AddOffice2" Then
                    CMDSQL = "update mgm set officenoadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskatofficeadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                 ElseIf LstReq.ListItems(W).SubItems(14) = "AddMobile1" Then
                    CMDSQL = "update mgm set mobilenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathpadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddMobile2" Then
                    CMDSQL = "update mgm set mobilenoadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathpadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddOther" Then
                    CMDSQL = "update mgm set req_nomor_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',status_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                Else
                    CMDSQL = "update mgm set req_nomor_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',status_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                End If
                
                CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(2)) + "'"
                M_OBJCONN.execute CMDSQL
                
                'Update Data Ke tabel LOg Telepon
                CMDSQL = "INSERT INTO tblrequestadditionalphone_log "
                CMDSQL = CMDSQL + "select * from tblrequestadditionalphone where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.execute CMDSQL
                
                'Update data log, tgl approve dan di approve oleh
                CMDSQL = "UPDATE tblrequestadditionalphone_log set tglapprove='"
                CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss") + "', approve_by='"
                CMDSQL = CMDSQL + MDIForm1.Text1.text + "' where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.execute CMDSQL
                
                'Hapus data di tabel tblrequestadditionalphone
                CMDSQL = "DELETE FROM tblrequestadditionalphone where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.execute CMDSQL
                
                '@@12-07-2011 Update status tanda request number
                CMDSQL = "UPDATE usertbl set f_req_number=null where userid in ("
                CMDSQL = CMDSQL + "select team from usertbl where userid='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(3) + "') "
                CMDSQL = CMDSQL + " or userid in (select userid  from usertbl where "
                CMDSQL = CMDSQL + " usertype='20' or usertype='25' or usertype='11') "
                M_OBJCONN.execute CMDSQL
                
                'Kasih tau agent
                CMDSQL = "INSERT INTO msgtbl (recipient,datetime,sender,msg) values ('"
                CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(3) + "','"
                CMDSQL = CMDSQL + CStr(Format(MDIForm1.TDBDate1.Value, "yyyymmdd")) + "','"
                CMDSQL = CMDSQL + CStr(MDIForm1.Text1.text) + "','"
                CMDSQL = CMDSQL + pesan + "')"
                M_OBJCONN.execute CMDSQL
                
                ' Masuk History 21 Juli 2014
                strket_hst = "Approve phone number : " & CStr(LstReq.ListItems(W).SubItems(11)) & "::" & CStr(LstReq.ListItems(W).SubItems(12))
                M_DATA.ADD_HISTORY LstReq.ListItems(W).SubItems(2), MDIForm1.TDBDate1.text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), LstReq.ListItems(W).SubItems(3), "COLLECTION", strket_hst, "", "", "", "", "", "", "", "", "", MDIForm1.Text1.text, "", "0", MDIForm1.txtdurasi.text, MDIForm1.txtuniqueid.text, "", cmb_topads.text, cmb_waiving.text, tdbwaiving.text, tdbamount_waiving.text, "0", "0"

'                If bAdd_phone = True Then
'next_step:
'                    err.CLEAR
'                    MsgBox "No Telp " & CStr(LstReq.ListItems(w).SubItems(11)) & " Pada req DC " & MDIForm1.Text1.Text & " Dibatalkan !", vbOKOnly + vbInformation, "INFO"
'
'                    ' COPY TO REJECT LIST
'                    M_OBJCONN.Execute "INSERT INTO tblrequestadditionalphone_reject SELECT * FROM tblrequestadditionalphone where id='" + LstReq.ListItems(w).Text + "'"
'
'                    ' Hapus data di tabel tblrequestadditionalphone
'                    cmdsql = "delete from tblrequestadditionalphone where id='"
'                    cmdsql = cmdsql + LstReq.ListItems(w).Text + "'"
'                    M_OBJCONN.Execute cmdsql
'
'                    LstReq.ListItems.CLEAR
'                    Call Isi_Req
'                    CmdApprove.Enabled = True
'                    CmdCekAll.Enabled = True
'                    CmdUnCekAll.Enabled = True
'                    CmdReject.Enabled = True
'
'                    Exit Sub
'                End If
        End If
    Next W
    
    MsgBox "Nomor Telepon berhasil di approve!", vbOKOnly + vbInformation, "Informasi"
    
    LstReq.ListItems.clear
    Call Isi_Req

    LstReqLog.ListItems.clear
    Call Isi_Req_log
    
    cmdapprove.Enabled = True
    CmdCekAll.Enabled = True
    CmdUncekAll.Enabled = True
    CmdReject.Enabled = True
End Sub

'Private Sub CmdApprove_Click()
'    Dim home1 As String
'    Dim home2 As String
'    Dim office1 As String
'    Dim office2 As String
'    Dim mobile1 As String
'    Dim mobile2 As String
'    Dim EcPhone As String
'    Dim W As Integer
'    Dim CMDSQL As String
'    Dim STRSQL As String
'    Dim m_objrs_waktu As ADODB.Recordset
'    Dim pesan As String
'    Dim string_pesan As String
'
'    If LstReq.ListItems.Count = 0 Then
'        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
'        Exit Sub
'    End If
'
'    'Ambil waktu dari server
'    CMDSQL = "select now() as waktu "
'    Set m_objrs_waktu = New ADODB.Recordset
'    m_objrs_waktu.CursorLocation = adUseClient
'    m_objrs_waktu.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    For W = 1 To LstReq.ListItems.Count
'        If LstReq.ListItems(W).Checked = True Then
'            home1 = IIf(IsNull(LstReq.ListItems(W).SubItems(4)), "", Trim(LstReq.ListItems(W).SubItems(4)))
'            home2 = IIf(IsNull(LstReq.ListItems(W).SubItems(5)), "", Trim(LstReq.ListItems(W).SubItems(5)))
'            office1 = IIf(IsNull(LstReq.ListItems(W).SubItems(6)), "", Trim(LstReq.ListItems(W).SubItems(6)))
'            office2 = IIf(IsNull(LstReq.ListItems(W).SubItems(7)), "", Trim(LstReq.ListItems(W).SubItems(7)))
'            mobile1 = IIf(IsNull(LstReq.ListItems(W).SubItems(8)), "", Trim(LstReq.ListItems(W).SubItems(8)))
'            mobile2 = IIf(IsNull(LstReq.ListItems(W).SubItems(9)), "", Trim(LstReq.ListItems(W).SubItems(9)))
'            EcPhone = IIf(IsNull(LstReq.ListItems(W).SubItems(10)), "", Trim(LstReq.ListItems(W).SubItems(10)))
'
'            STRSQL = ""
'            pesan = ""
'
'            If home1 <> "" Then
'                If STRSQL = "" Then
'                    STRSQL = " homenoadd1='" + home1 + "' "
'                    pesan = "Additional Home 1"
'                Else
'                    STRSQL = STRSQL + ", homenoadd1='" + home1 + "' "
'                    pesan = pesan + ", Additional Home 1"
'                End If
'            End If
'
'            If home2 <> "" Then
'                If STRSQL = "" Then
'                    STRSQL = " homenoadd2='" + home2 + "' "
'                    pesan = "Additional Home 2"
'                Else
'                    STRSQL = STRSQL + ", homenoadd2='" + home2 + "' "
'                    pesan = pesan + ", Additional Home 2"
'                End If
'            End If
'
'            If office1 <> "" Then
'                If STRSQL = "" Then
'                    STRSQL = " officenoadd1='" + office1 + "' "
'                    pesan = "Additional Office 1"
'                Else
'                    STRSQL = STRSQL + ", officenoadd1='" + office1 + "' "
'                    pesan = pesan + ", Additional Office 1"
'                End If
'            End If
'
'            If office2 <> "" Then
'                If STRSQL = "" Then
'                    STRSQL = " officenoadd2='" + office2 + "' "
'                    pesan = "Additional Office 2"
'
'                Else
'                    STRSQL = STRSQL + ", officenoadd2='" + office2 + "' "
'                    pesan = pesan + ", Additional Office 2"
'                End If
'            End If
'
'
'            If mobile1 <> "" Then
'                If STRSQL = "" Then
'                    STRSQL = " mobilenoadd1='" + mobile1 + "' "
'                    pesan = "Additional Mobile 1"
'                Else
'                    STRSQL = STRSQL + ", mobilenoadd1='" + mobile1 + "' "
'                    pesan = pesan + ", Additional Mobile 1"
'                End If
'            End If
'
'            If mobile2 <> "" Then
'                If STRSQL = "" Then
'                    STRSQL = " mobilenoadd2='" + mobile2 + "' "
'                    pesan = "Additional Mobile 2"
'                Else
'                    STRSQL = STRSQL + ", mobilenoadd2='" + mobile2 + "' "
'                    pesan = pesan + ", Additional Mobile 2"
'                End If
'            End If
'
'             If EcPhone <> "" Then
'                If STRSQL = "" Then
'                    STRSQL = " ec_telp='" + EcPhone + "' "
'                    pesan = "Ec Phone"
'                Else
'                    STRSQL = STRSQL + ", ec_telp='" + EcPhone + "' "
'                    pesan = pesan + ", Ec Phone"
'                End If
'            End If
'
'            If STRSQL = "" Then
'                '@@ 16-06-2011 ini jika telepon kosong maka langsung hapus
'                CMDSQL = "delete from tblrequestadditionalphone where id='"
'                CMDSQL = CMDSQL + LstReq.ListItems(W).Text + "'"
'                M_OBJCONN.Execute CMDSQL
'                GoTo lanjut
'            End If
'
'            'Pindahkan data dari tabel tblrequestadditionalphone ke tblrequestadditionalphone_log
'            CMDSQL = "insert into tblrequestadditionalphone_log "
'            CMDSQL = CMDSQL + "select * from tblrequestadditionalphone where id='"
'            CMDSQL = CMDSQL + LstReq.ListItems(W).Text + "'"
'            M_OBJCONN.Execute CMDSQL
'
'
'            'Update ke mgm
'            CMDSQL = "update mgm set " + STRSQL
'            CMDSQL = CMDSQL + " where custid='"
'            CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(2) + "'"
'
'            M_OBJCONN.Execute CMDSQL
'
'            'Update data log, tgl approve dan di approve oleh
'            CMDSQL = "update tblrequestadditionalphone_log set tglapprove='"
'            CMDSQL = CMDSQL + Format(m_objrs_waktu(0), "yyyy-mm-dd hh:mm:ss") + "', approve_by='"
'            CMDSQL = CMDSQL + MDIForm1.Text1.Text + "' where id='"
'            CMDSQL = CMDSQL + LstReq.ListItems(W).Text + "'"
'            M_OBJCONN.Execute CMDSQL
'
'            'Hapus data di tabel tblrequestadditionalphone
'            CMDSQL = "delete from tblrequestadditionalphone where id='"
'            CMDSQL = CMDSQL + LstReq.ListItems(W).Text + "'"
'            M_OBJCONN.Execute CMDSQL
'
'            '@@12-07-2011 Update status tanda request number
'            CMDSQL = "update usertbl set f_req_number=null where userid in ("
'            CMDSQL = CMDSQL + "select team from usertbl where userid='"
'            CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(3) + "') "
'            CMDSQL = CMDSQL + " or userid in (select userid  from usertbl where "
'            CMDSQL = CMDSQL + " usertype='20' or usertype='25' or usertype='11') "
'            M_OBJCONN.Execute CMDSQL
'
'
'            string_pesan = string_pesan + Chr(13) + "Custid:" + LstReq.ListItems(W).SubItems(2)
'            string_pesan = string_pesan + " berhasil diupdate: " + pesan + Chr(13)
'        End If
'lanjut:
'    Next W
'
'    MsgBox string_pesan, vbOKOnly + vbInformation, "Informasi"
'
'    'Update status tanda request number
''    CMDSQL = "update usertbl set f_req_number=null where userid='"
''    CMDSQL = CMDSQL + MDIForm1.Text1.Text + "' or usertype='20' or usertype='25' or usertype='11' or usertype='6'"
''    M_OBJCONN.Execute CMDSQL
'
'    LstReq.ListItems.CLEAR
'    Set m_objrs_waktu = Nothing
'
'    Call Isi_Req
'
'    LstReqLog.ListItems.CLEAR
'    Call Isi_Req_log
'End Sub

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LstReq.ListItems.Count
        LstReq.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdReject_Click()
    Dim W As Integer
    Dim CMDSQL As String
    Dim pesan As String
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    pesan = MsgBox("Yakin data mau dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If pesan = vbYes Then
        For W = 1 To LstReq.ListItems.Count
            If LstReq.ListItems(W).Checked = True Then
               CMDSQL = "delete from tblrequestadditionalphone where id='"
               CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
               M_OBJCONN.execute CMDSQL
               
               '@@12-07-2011 Update status tanda request number
                CMDSQL = "update usertbl set f_req_number=null where userid in ("
                CMDSQL = CMDSQL + "select team from usertbl where userid='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(3) + "') "
                CMDSQL = CMDSQL + " or userid in (select userid  from usertbl where "
                CMDSQL = CMDSQL + " usertype='20' or usertype='25' or usertype='11') "
               M_OBJCONN.execute CMDSQL
            End If
        Next W
        
        'Update status tanda request number
'        CMDSQL = "update usertbl set f_req_number=null where userid='"
'        CMDSQL = CMDSQL + MDIForm1.Text1.Text + "' or usertype='20' or usertype='25' or usertype='11' "
'        M_OBJCONN.Execute CMDSQL
        
        MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
        LstReq.ListItems.clear
        Call Isi_Req
    End If
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LstReq.ListItems.Count
        LstReq.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Command1_Click()
    If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.text) = "MANAGER" Or UCase(MDIForm1.Text2.text) = "ADMINISTRATOR" Then
        Form_addphone_reject.Show 1
    End If
End Sub

Private Sub Form_Load()
    Call HeaderLstReq
    Call Isi_Req_log
    
    Call HeaderReq
    Call Isi_Req
End Sub

Private Sub Isi_Req()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        CMDSQL = "select * from tblrequestadditionalphone where agent in "
        CMDSQL = CMDSQL + " (select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.Text1.text + "') "
        CMDSQL = CMDSQL + " order by tglreq desc "
    Else
        CMDSQL = "select * from tblrequestadditionalphone order by tglreq desc"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    While Not M_objrs.EOF
        Set ListItem = LstReq.ListItems.ADD(, , M_objrs("id"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("tglreq")), "", M_objrs("tglreq"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("home1")), "", M_objrs("home1"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("home2")), "", M_objrs("home2"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("office1")), "", M_objrs("office1"))
            ListItem.SubItems(7) = IIf(IsNull(M_objrs("office2")), "", M_objrs("office2"))
            ListItem.SubItems(8) = IIf(IsNull(M_objrs("mobile1")), "", M_objrs("mobile1"))
            ListItem.SubItems(9) = IIf(IsNull(M_objrs("mobile2")), "", M_objrs("mobile2"))
            ListItem.SubItems(10) = IIf(IsNull(M_objrs("ecphone")), "", M_objrs("ecphone"))
            ListItem.SubItems(11) = IIf(IsNull(M_objrs("request_number")), "", M_objrs("request_number"))
            ListItem.SubItems(12) = IIf(IsNull(M_objrs("kategori")), "", M_objrs("kategori"))
            ListItem.SubItems(13) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
            ListItem.SubItems(14) = IIf(IsNull(M_objrs("jenis")), "", M_objrs("jenis"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
    
    TxtReq.text = LstReq.ListItems.Count
End Sub


Private Sub Isi_Req_log()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        CMDSQL = "select * from tblrequestadditionalphone_log where agent in "
        CMDSQL = CMDSQL + " (select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.Text1.text + "') "
        CMDSQL = CMDSQL + " order by tglreq desc limit 100"
    Else
        CMDSQL = "select * from tblrequestadditionalphone_log order by tglreq desc limit 100"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    
    While Not M_objrs.EOF
        Set ListItem = LstReqLog.ListItems.ADD(, , M_objrs("tglreq"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("home1")), "", M_objrs("home1"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("home2")), "", M_objrs("home2"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("office1")), "", M_objrs("office1"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("office2")), "", M_objrs("office2"))
            ListItem.SubItems(7) = IIf(IsNull(M_objrs("mobile1")), "", M_objrs("mobile1"))
            ListItem.SubItems(8) = IIf(IsNull(M_objrs("mobile2")), "", M_objrs("mobile2"))
            ListItem.SubItems(9) = IIf(IsNull(M_objrs("ecphone")), "", M_objrs("ecphone"))
            ListItem.SubItems(10) = IIf(IsNull(M_objrs("tglapprove")), "", M_objrs("tglapprove"))
            ListItem.SubItems(11) = IIf(IsNull(M_objrs("approve_by")), "", M_objrs("approve_by"))
            ListItem.SubItems(12) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
        M_objrs.MoveNext
    Wend
    
    TxtReqLog.text = LstReqLog.ListItems.Count
End Sub


