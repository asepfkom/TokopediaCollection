VERSION 5.00
Begin VB.Form FrmClaimAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Claim Account"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3180
      TabIndex        =   7
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton CmdProses 
      Caption         =   "Proses"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2400
      Width           =   1035
   End
   Begin VB.TextBox TxtAlasanClaim 
      Height          =   1605
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox TxtNama 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   420
      Width           =   3015
   End
   Begin VB.TextBox Txtcustid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Alasan claim:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nama:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Custid:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmClaimAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdProses_Click()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim a, Pesan, RemarksClaim As String
    
    If TxtAlasanClaim.Text = "" Then
        MsgBox "Alasan claim tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin akan memproses claim account?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    On Error GoTo salah
    TxtAlasanClaim.Enabled = False
    DoEvents
    
    
    'Update ke mgm pindahkan ke account claim
    Cmdsql = "update mgm set agent='CLAIM',user_claim='"
    Cmdsql = Cmdsql + MDIForm1.Text1.Text + "', waktu_claim=now(),alasan_claim='"
    Cmdsql = Cmdsql + Replace(TxtAlasanClaim.Text, "'", "") + "' "
    Cmdsql = Cmdsql + " where custid='"
    Cmdsql = Cmdsql + CStr(TxtCustid.Text) + "'"
    M_OBJCONN.Execute Cmdsql
    
    'Catet di history nih...
    RemarksClaim = "Agent : " & MDIForm1.Text1.Text & "=> Telah Melakukan Claim pada account ini <="
    RemarksClaim = RemarksClaim & Replace(TxtAlasanClaim.Text, "'", "")
    
    Cmdsql = "insert into mgm_hst (custid,agent,hst,tgl,user_log) values ('"
    Cmdsql = Cmdsql + CStr(TxtCustid.Text) + "','"
    Cmdsql = Cmdsql & FrmCC_Colection.lblaoc.Caption + "','"
    Cmdsql = Cmdsql & RemarksClaim & "',now(),'"
    Cmdsql = Cmdsql & MDIForm1.Text1.Text & "')"
    M_OBJCONN.Execute Cmdsql
    
    ' UPDATED 22 MEI 2013 - IZUDDIN
    Cmdsql = "insert into tbllog_claim_aksesall (custid,agent,agentlama,tgl_claim) values ('"
    Cmdsql = Cmdsql + CStr(TxtCustid.Text) + "','"
    Cmdsql = Cmdsql & MDIForm1.Text1.Text + "','"
    Cmdsql = Cmdsql & FrmCC_Colection.lbl_agentlama.Caption + "', "
    Cmdsql = Cmdsql & "now())"
    M_OBJCONN.Execute Cmdsql
    
    'Kirim pesan ke semua agent yang ada di distribusi
    Pesan = "Pesan ini dibuat otomatis oleh system " & vbCrLf
    Pesan = Pesan & "========================================" & vbCrLf
    Pesan = Pesan & "Agent : " & MDIForm1.Text1.Text & vbCrLf
    Pesan = Pesan & "Telah melakukan Claim untuk account : " & vbCrLf & vbCrLf
    Pesan = Pesan & "Custid :" & TxtCustid.Text & vbCrLf
    Pesan = Pesan & "Nama :" & TxtNama.Text & vbCrLf & vbCrLf
    Pesan = Pesan & "Alasan untuk claim: " & vbCrLf
    Pesan = Pesan & Replace(TxtAlasanClaim.Text, "'", "")
    
    '--1. Kirim ke TL dan SPV
    'Cmdsql = "select * from usertbl where usertype in ('11','6','20','25') "
    '@@20022013 Jika agent mengclaim account, pesan ga usah ditampilkan ke spv
    Cmdsql = "select * from usertbl where usertype in ('6','20','25') "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            Cmdsql = "insert into msgtbl "
            Cmdsql = Cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            Cmdsql = Cmdsql + M_Objrs("userid") + "','"
            Cmdsql = Cmdsql + Format(Now(), "yyyymmdd") + "','"
            Cmdsql = Cmdsql + MDIForm1.Text1.Text + "','"
            Cmdsql = Cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            Cmdsql = Cmdsql + Pesan + "')"
            M_OBJCONN.Execute Cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
    
    '--2. Kirim juga agent yang ada di tabel distribusi yang ikut mengcollect custid ini
'    Cmdsql = "select agent from tbl_distribusi_account where custid='"
'    Cmdsql = Cmdsql & CStr(TxtCustid.Text) & "'"
    Cmdsql = "SELECT a.*,b.userid as agent FROM tbl_cust_aksesall a,usertbl b WHERE a.kd_profile=b.profile_akses_all "
    Cmdsql = Cmdsql & " AND a.custid='" & CStr(TxtCustid.Text) & "'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            Cmdsql = "insert into msgtbl "
            Cmdsql = Cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            Cmdsql = Cmdsql + M_Objrs("agent") + "','"
            Cmdsql = Cmdsql + Format(Now(), "yyyymmdd") + "','"
            Cmdsql = Cmdsql + MDIForm1.Text1.Text + "','"
            Cmdsql = Cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            Cmdsql = Cmdsql + Pesan + "')"
            M_OBJCONN.Execute Cmdsql
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
    
    'Hapus data dari tbl_distribusi_account
'    Cmdsql = "delete from tbl_distribusi_account where custid='"
    Cmdsql = "delete from tbl_cust_aksesall where custid='"
    Cmdsql = Cmdsql & CStr(TxtCustid.Text) & "'"
    M_OBJCONN.Execute Cmdsql
    
    MsgBox "Proses claim sudah dikirim! Jika di ACC sistem akan memberitahukan kepada anda!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
    Exit Sub
salah:
    MsgBox "Mohon maaf ada kesalahan: " & err.Description
    
End Sub
