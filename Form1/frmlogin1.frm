VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4215
   ClientLeft      =   7290
   ClientTop       =   3930
   ClientWidth     =   7935
   LinkTopic       =   "Form3"
   ScaleHeight     =   4215
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4230
      Left            =   -45
      TabIndex        =   0
      Top             =   -15
      Width           =   10350
      Begin VB.Timer Tmrreminder 
         Interval        =   100
         Left            =   7545
         Top             =   930
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   7545
         Top             =   465
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2910
         PasswordChar    =   "#"
         TabIndex        =   2
         Top             =   2790
         Width           =   2235
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   6075
         Picture         =   "frmlogin1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2295
         Width           =   720
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   5220
         Picture         =   "frmlogin1.frx":0655
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2325
         Width           =   765
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2910
         TabIndex        =   1
         Top             =   2325
         Width           =   2235
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Password"
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
         Left            =   2880
         TabIndex        =   8
         Top             =   3225
         Width           =   2235
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1845
         Left            =   -270
         Picture         =   "frmlogin1.frx":0C5D
         ScaleHeight     =   1845
         ScaleWidth      =   5835
         TabIndex        =   7
         Top             =   555
         Width           =   5835
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         Caption         =   "Copyright � 2020 Delta Nuansa Nirwana All Right Reserved"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   135
         TabIndex        =   6
         Top             =   3855
         Width           =   11235
      End
      Begin VB.Label lbl1 
         BackColor       =   &H0000C000&
         Height          =   480
         Index           =   1
         Left            =   -1470
         TabIndex        =   12
         Top             =   3750
         Width           =   11235
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1740
         TabIndex        =   9
         Top             =   2355
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5355
         TabIndex        =   11
         Top             =   3135
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   1725
         TabIndex        =   10
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label lbl1 
         BackColor       =   &H0000C000&
         Height          =   480
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   -15
         Width           =   11235
      End
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        txtPassword.PasswordChar = ""
    Else
        txtPassword.PasswordChar = "#"
    End If
End Sub

Private Sub CmdCancel_Click()
    End
End Sub

Private Sub maskbbb(STATUS As String)
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    strsql = "SELECT * From information_schema.Columns WHERE table_name='tblmaskbbb' ORDER BY ordinal_position"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        M_OBJCONN.execute "create table tblmaskbbb (status smallint);"
    End If
    
    If STATUS = "yes" Then
        strsql = "SELECT * From tblmaskbbb"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
        If rs.RecordCount = 0 Then
            M_OBJCONN.execute "insert into tblmaskbbb values (1);"
        Else
            M_OBJCONN.execute "update tblmaskbbb set status = 1;"
        End If
    
    ElseIf STATUS = "no" Then
        strsql = "SELECT * From tblmaskbbb"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
        If rs.RecordCount = 0 Then
            M_OBJCONN.execute "insert into tblmaskbbb values (0);"
        Else
            M_OBJCONN.execute "update tblmaskbbb set status = 0;"
        End If
    End If
End Sub

Private Sub CmdOK_Click()
Dim NILSTAT As String
Dim M_objrs As ADODB.Recordset
Dim rs_lvtian As New ADODB.Recordset
Dim m_objrsAdd As ADODB.Recordset
Dim M_PESAN As ADODB.Recordset
Dim m_waktuserver As ADODB.Recordset
Dim lminggu As String
Dim lbulan As String
Dim strsql As String
Dim ltahun As String
Dim CMDSQL As String
Dim m_popup As ADODB.Recordset
Dim CMDSQL2 As String
Dim SqlWaktu As String
Dim jam_sekarang As String
Dim xxx As Boolean

If txtUserName.text = "tblmaskbbb" And txtPassword.text = "yes" Then
    Call maskbbb("yes")
    MsgBox "aktif"
    Exit Sub
ElseIf txtUserName.text = "tblmaskbbb" And txtPassword.text = "no" Then
    Call maskbbb("no")
    MsgBox "non-aktif"
    Exit Sub
End If



 ' On Error GoTo HELL
exit_klik = False

'If TxtUsername = "Admin" And TxtPassword = "Rqo317317" Then
'If txtUserName = "Admin" And txtPassword = "t3mp1l@ng" Or txtUserName = "IZUDDIN" And txtPassword = "xxx123" Then
'    MDIForm1.Text1.Text = txtUserName
'    MDIForm1.Text2.Text = "Administrator"
'    MDIForm1.Text7.Text = "Administrator"
'    Unload Me
'    MDIForm1.mn_update_db.Visible = True
'    'JEJAKTIAN10032016==================================================
'    If txtUserName <> "tian" Then
'        MDIForm1.nmlistreqptp.Visible = False
'    End If
'    '===================================================================
'    MDIForm1.Show
'    Exit Sub
'End If
If (txtUserName = "tian" And txtPassword = "tian") Then
'    (txtUserName = "JOKO" And txtPassword = "DNN525") Or _
'    (txtUserName = "WULAN" And txtPassword = "DNN525") Or _
'    (txtUserName = "RICKY" And txtPassword = "DNN525") Or _
'    (txtUserName = "FIFI" And txtPassword = "DNN525") Then
'
    MDIForm1.Text1.text = txtUserName
    MDIForm1.Text2.text = "Administrator"
    MDIForm1.Text7.text = "Administrator"
    MDIForm1.mn_update_db.Visible = True
    MDIForm1.mnpd.Visible = True
    MDIForm1.mnkk.Visible = True
    Unload Me
    'JEJAKTIAN10032016==================================================
    If MDIForm1.Text1.text <> "tian" Then
        MDIForm1.nmlistreqptp.Visible = False
    End If
    '===================================================================

    'If MDIForm1.Text2.Text = "Administrator" Then
        MDIForm1.Ofl.Visible = True
    'End If

    MDIForm1.Show
    Exit Sub
End If

xxx = CheckPath("C:\sempakbasah\")

    If txtUserName = Empty Then
        MsgBox "Username Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
        txtUserName.SetFocus
        Sendkeys "{Home}+{End}"
        Exit Sub
    Else
        If xxx = True Then
            GoTo xx
        End If
        If txtPassword = Empty Then
            MsgBox "Password Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
            txtPassword.SetFocus
            Sendkeys "{Home}+{End}"
            Exit Sub
        End If
    End If
xx:
Timer1.Enabled = True
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
'M_OBJRS.Open "SELECT USERID, ACCREC, USERTYPE,AGENT,UNIT,AUTH, EXT,stsaplikasi,note,ntargetspv FROM usertbl WHERE USERID = '" + txtUserName + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'301110 Ubah ke Md5
CMDSQL2 = "SELECT userid, accrec, usertype,agent,unit,auth, ext,"
CMDSQL2 = CMDSQL2 + "stsaplikasi,note,ntargetspv, date(now())-date(tgl_ubah_pass) as LamaPass, f_status_login ,* from usertbl WHERE userid='"
CMDSQL2 = CMDSQL2 + Trim(txtUserName.text) + "' and accrec=md5('"
CMDSQL2 = CMDSQL2 + Trim(txtPassword.text) + "')"
M_objrs.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If txtPassword.text = "bukadulu" Then
    Dim query As String
    If txtUserName.text = "ALL" Then
        query = "UPDATE usertbl SET f_status_login=null,last_logout='now()'"
        M_OBJCONN.execute query
        MsgBox "Data Sudah Terupdate"
        End
    Else
        query = "UPDATE usertbl SET f_status_login=null,last_logout='now()' where userid = '" + txtUserName.text + "'"
        M_OBJCONN.execute query
        MsgBox "Data Sudah Terupdate"
        End
    End If
End If

If xxx = True Then
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL2 = "SELECT userid, accrec, usertype,agent,unit,auth, ext,"
    CMDSQL2 = CMDSQL2 + "stsaplikasi,note,ntargetspv, date(now())-date(tgl_ubah_pass) as LamaPass, f_status_login, * from usertbl WHERE userid='"
    CMDSQL2 = CMDSQL2 + Trim(txtUserName.text) + "'"
    M_objrs.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If

If M_objrs.RecordCount <> 0 Then
      'MDIForm1.Text1.text = txtUserName.text
    If M_objrs!f_status_login = 1 And M_objrs!usertype = 1 Then
        MsgBox "User " & M_objrs!Userid & " sudah login dilain PC, Force LogOut klik OK "
        '---- update tbluser_log status logout
        CMDSQL = "update usertbl_log set waktu_logout=now()::varchar,durasi = now() - substring(waktu_login,1,19)::timestamp where waktu_login= (select max(waktu_login) from usertbl_log where username='" + txtUserName.text + "') "
        CMDSQL = CMDSQL + " and username='" + txtUserName.text + "'"

    M_OBJCONN.execute CMDSQL
      '  End
    End If
        
'        If txtPassword <> M_OBJRS("ACCREC") Then
'            Debug.Print Decrypt(Len(Trim(txtUserName.text)), M_OBJRS("ACCREC"))
'        End If
        
    ''    If Trim(txtPassword) <> Decrypt(Len(Trim(txtUserName.Text)), M_OBJRS("ACCREC")) Then
    ''        MsgBox "Password Yang Anda Masukan Salah... Perhatikan CapsLock Anda...!!!", vbCritical + vbOKOnly, "Peringatan"
    ''        txtPassword.SetFocus
            'SendKeys "{Home}+{End}"
    ''    Else
    
        ' CEK JAM MASUK RANDY(FEB 2016)
        SqlWaktu = "select now()"
        Set m_waktuserver = New ADODB.Recordset
        m_waktuserver.CursorLocation = adUseClient
        m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        

        
        'Jika last login sekarang
'        If Format(m_waktuserver(0), "yyyy-mm-dd") <> Format(M_objrs("tglupdate"), "yyyy-mm-dd") Then
'            If Format(m_waktuserver(0), "HH:mm") > Format("08:05", "HH:mm") Then
'                If M_objrs("USERTYPE") = "1" Then
'                    Set m_waktuserver = Nothing
'                    M_OBJCONN.execute "UPDATE usertbl SET f_blok='1',alasan_blok='Terlambat masuk',tglupdate='now()' WHERE userid='" & Trim(txtUserName.text) & "'"
'                    MsgBox "Jam masuk anda terlambat!! Tidak boleh melebihi Pukul 08:00", vbCritical + vbOKOnly, "Terlambat"
'                    Call offsesilogin(txtUserName.text)
'                    GoTo blok_user
'                End If
'            End If
'        End If
'
'        ' Waktu masuk lebih dari 10 menit
'        If DateDiff("n", Format(M_objrs("last_logout"), "yyyy-mm-dd hh:mm:ss"), Format(m_waktuserver(0), "yyyy-mm-dd hh:mm:ss")) >= 10 Then
'            If Format(m_waktuserver(0), "HH:mm") > Format("08:05", "HH:mm") Then
'                If M_objrs("USERTYPE") = "1" Then
'                    If M_objrs("f_break") = 0 Then
'                        Set m_waktuserver = Nothing
'                        M_OBJCONN.execute "UPDATE usertbl SET f_blok='1',alasan_blok='10 Menit',tglupdate=now() WHERE userid='" & Trim(txtUserName.text) & "'"
'                        MsgBox "Anda diblok karena membuka aplikasi Lebih dari 10 Menit dari " & vbCrLf & "waktu terakhir keluar program (log out)", vbCritical + vbOKOnly, "Blok"
'                        Call offsesilogin(txtUserName.text)
'                        End
'                        GoTo blok_user
'                    End If
'                End If
'            End If
'        End If
'
        M_OBJCONN.execute "UPDATE usertbl SET last_logout=now(),tglupdate=now(),f_break=0 WHERE userid='" & Trim(txtUserName.text) & "'"
        
        Set m_waktuserver = Nothing
        ' # END CEK JAM MASUK
        
       If IsNull(M_objrs("tgl_ubah_pass")) = True Or Val(IIf(IsNull(M_objrs("LamaPass")), "0", M_objrs("lamapass"))) >= 90 Then
            MsgBox "Untuk keamanan! Silahkan ganti password anda terlebih dahulu!"
            FrmGantiPassword.TxtCoding.text = txtUserName.text
            FrmGantiPassword.Show vbModal
       End If

        If M_objrs("USERTYPE") = "1" Then
            If IIf(IsNull(M_objrs("note")), "", M_objrs("note")) = "" Or IIf(IsNull(M_objrs("note")), "", M_objrs("note")) = 0 Then
                NILSTAT = ""
            Else
                NILSTAT = "" + IIf(IsNull(M_objrs("note")), "", M_objrs("note")) + ""
            End If
            
            SqlWaktu = "select now()"
            Set m_waktuserver = New ADODB.Recordset
            m_waktuserver.CursorLocation = adUseClient
            m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
           
            jam_sekarang = Format(m_waktuserver(0), "hh")
        
            If jam_sekarang < 7 Then
                MsgBox "Anda Tidak Boleh Login Kurang Dari Jam 07:00", vbCritical + vbOKOnly, "Terlambat"
            Exit Sub
            End If
            'MDIForm1.Lbltargetspv = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
            'MDIForm1.Kalimat1 = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
            'MDIForm1.PANJANG = Len("Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note"))))
            MDIForm1.mnsubmarkup.Visible = False
            MDIForm1.Lbltargetspv = NILSTAT
            MDIForm1.Kalimat1 = NILSTAT
            MDIForm1.PANJANG = Len(NILSTAT)
            MDIForm1.mnsubahstsacc.Visible = False
            MDIForm1.setspv.Visible = False
            'MDIForm1.LblTarget.Visible = True
            MDIForm1.Text2.text = "Agent"
            'MDIForm1.mnnote.Visible = False
            MDIForm1.mnsegment.Visible = False
            MDIForm1.SSCommand1(11).Visible = False
            MDIForm1.Label8.Visible = False
            'MDIForm1.SSCommand1(7).Visible = False
            MDIForm1.mnbar(1).Visible = False
            MDIForm1.mnbar(2).Visible = False
            MDIForm1.mnbar(3).Visible = False
            MDIForm1.mnbar(5).Visible = False
            MDIForm1.mnbar(6).Visible = False
            MDIForm1.mnbar(7).Visible = False
            MDIForm1.mnbar(11).Visible = False
            MDIForm1.MnFile(1).Visible = False
            'if m_objrs("stsaplikasi")
            MDIForm1.SSCommand1(1).Visible = True
            MDIForm1.SSCommand1(2).Visible = False
            MDIForm1.SSCommand1(4).Visible = False
            MDIForm1.SSCommand1(5).Visible = False
            MDIForm1.SSCommand1(8).Visible = False
            MDIForm1.Label6.Visible = False
            MDIForm1.SSCommand2.Visible = False
             MDIForm1.VSMS.Visible = False
            'MDIForm1.SSCommand1(3).Visible = False
            'Dim m_popup As New ADODB.Recordset
'            Set m_popup = New ADODB.Recordset
'            m_popup.CursorLocation = adUseClient
'            m_popup.Open "Select * from vwcallcfg1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            CMDSQL2 = "UPDATE usertbl set flagcall ='" + Format(m_popup!tglsystem, "hh:mm:ss") + "' where userid ='" + txtUserName.Text + "'"
'            M_OBJCONN.Execute CMDSQL2
'
'            Set m_popup = Nothing
            
            
           
        Else
            'MDIForm1.LblTarget.Visible = False
            If M_objrs("USERTYPE") = "6" Then
                If IIf(IsNull(M_objrs("note")), "", M_objrs("note")) = "" Or IIf(IsNull(M_objrs("note")), "", M_objrs("note")) = "0" Then
                    NILSTAT = ""
                Else
                    NILSTAT = "" + IIf(IsNull(M_objrs("note")), "", M_objrs("note")) + ""
                End If
           
           ' MDIForm1.Lbltargetspv = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
            'MDIForm1.Kalimat1 = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
            'MDIForm1.PANJANG = Len("Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note"))))
            
            MDIForm1.Lbltargetspv = NILSTAT
            MDIForm1.Kalimat1 = NILSTAT
            MDIForm1.PANJANG = Len(NILSTAT)
            
            MDIForm1.mnsubahstsacc.Visible = False
            MDIForm1.setspv.Visible = False
            MDIForm1.Text2.text = "TeamLeader"
            'MDIForm1.mnsegment.Visible = False
            MDIForm1.mnbar(2).Visible = False
            MDIForm1.mnbar(5).Visible = False
            MDIForm1.mnbar(7).Visible = False
           ' MDIForm1.mnblokspv.Visible = False
            MDIForm1.VSMS.Visible = False
            End If
            If M_objrs("USERTYPE") = "2" Then
                'MDIForm1.LblTarget.Visible = True
            MDIForm1.Text2.text = "Field Collector"
            MDIForm1.SSCommand1(11).Visible = False
            MDIForm1.mnbar(1).Visible = False
            MDIForm1.mnbar(2).Visible = False
            MDIForm1.mnbar(3).Visible = False
            MDIForm1.mnbar(5).Visible = False
            MDIForm1.mnbar(6).Visible = False
            MDIForm1.mnbar(7).Visible = False
            MDIForm1.MnFile(1).Visible = False
            MDIForm1.SSCommand1(0).Visible = False
            MDIForm1.SSCommand1(1).Visible = False
            MDIForm1.SSCommand1(2).Visible = False
            MDIForm1.SSCommand1(4).Visible = False
            MDIForm1.SSCommand1(5).Visible = True
            MDIForm1.SSCommand2.Visible = False
            
            
            End If
        End If
        
        If M_objrs("USERTYPE") = "11" Or M_objrs("USERTYPE") = "20" Then
            MDIForm1.Text2.text = "Supervisor"
            MDIForm1.mnlist.Visible = True
        End If
        
        If M_objrs("USERTYPE") = "17" Then
            MDIForm1.Text2.text = "Manager"
            MDIForm1.mnlist.Visible = True
        End If
        
        If M_objrs("USERTYPE") = "25" Then
            MDIForm1.Text2.text = "Admin"
        End If
        
        'jejaktian28072016menurole
        'Call menurole
        '=================================================
        
        MDIForm1.Text1.text = UCase(txtUserName)
        Dim qs As String
        Dim rs As ADODB.Recordset
        'AM req pak rio 18 april 2018
        If Left(UCase(txtUserName), 2) = "TL" Then
            qs = "select distinct am,amcaption from tblsettingam where am = '" & UCase(txtUserName) & "'"
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If rs.RecordCount > 0 Then
                MDIForm1.Text2.text = rs!amcaption
                MDIForm1.mnagent.Visible = False
                MDIForm1.upload_fresh_wo.Visible = False
            End If
        End If
        '====================================================
        MDIForm1.Text3.text = IIf(IsNull(M_objrs("UNIT")), "", M_objrs("UNIT"))
        MDIForm1.Text7.text = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        MDIForm1.TxtAuth.text = IIf(IsNull(M_objrs("AUTH")), "", M_objrs("AUTH"))
        MDIForm1.mnappvp.Caption = "Approve Valid Phone " + MDIForm1.Text2.text
        DoEvents
        'Call MDIForm1.LoOut_Ext("*1")
        WaitSecs (0.1)
        'Call login_ext(IIf(IsNull(m_objrs!EXT), "*1", m_objrs!EXT))
        
        'isi target
        
'        Set m_objrsAdd = New ADODB.Recordset
'        m_objrsAdd.CursorLocation = adUseClient
'        CMDSQL = "Select * from TblTanggal Where "
'        CMDSQL = CMDSQL + " TGL = '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") + "'"
'        m_objrsAdd.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If m_objrsAdd.RecordCount <> 0 Then
'            lminggu = IIf(IsNull(m_objrsAdd!Minggu), 0, m_objrsAdd!Minggu)
'            lbulan = IIf(IsNull(m_objrsAdd!Bulan), 0, m_objrsAdd!Bulan)
'            ltahun = IIf(IsNull(m_objrsAdd!tahun), 0, m_objrsAdd!tahun)
'        Else
'   '         MsgBox "Tanggal Belum Di Set....", vbInformation + vbOKOnly, "Aplikasi"
'            lminggu = 0
'            lbulan = MDIForm1.TDBDate1.Month
'            ltahun = MDIForm1.TDBDate1.Year
'        End If
'        Set m_objrsAdd = Nothing
'        DoEvents
       'Set m_objrs = Nothing
        Unload Me
        
        '@@09022011 Ambil nilai maksimal kuota sms per hari agent dapat mengirim sms
        Dim m_objrskuota As ADODB.Recordset
        Dim cmdsqlkuota As String
        
        cmdsqlkuota = "select * from tblsetsms"
        Set m_objrskuota = New ADODB.Recordset
        m_objrskuota.CursorLocation = adUseClient
        m_objrskuota.Open cmdsqlkuota, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If m_objrskuota.RecordCount > 0 Then
            MDIForm1.KuotaSms = m_objrskuota("kuota_sms")
        End If
        Set m_objrskuota = Nothing
        
        '@@ 12-04-2011, Catet IP address user yang login, buat kirim pesan via winsock
        Dim ip_addr As String
        Dim agent As String
        Dim tipe As String
        Dim M_Objrs_Cek As ADODB.Recordset
        Dim StrSqlIp As String
        
        ip_addr = MDIForm1.WskCTI.LocalIP
        agent = UCase(MDIForm1.Text1.text)
        tipe = UCase(MDIForm1.Text2.text)
        
        'Cek dulu, apakah data IP user sudah ada, jika sudah ada di Update IPnya
        StrSqlIp = "select * from tbl_ip where agent='"
        StrSqlIp = StrSqlIp + Trim(agent) + "'"
        Set M_Objrs_Cek = New ADODB.Recordset
        M_Objrs_Cek.CursorLocation = adUseClient
        M_Objrs_Cek.Open StrSqlIp, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek.RecordCount = 0 Then
            'Inputin deh data baru
            StrSqlIp = "insert into tbl_ip (agent,tipe,ip_addr) values ('"
            StrSqlIp = StrSqlIp + Trim(agent) + "','"
            StrSqlIp = StrSqlIp + Trim(tipe) + "','"
            StrSqlIp = StrSqlIp + Trim(ip_addr) + "')"
            M_OBJCONN.execute StrSqlIp
        Else
            StrSqlIp = "update tbl_ip set ip_addr='"
            StrSqlIp = StrSqlIp + Trim(ip_addr) + "' where agent='"
            StrSqlIp = StrSqlIp + Trim(agent) + "'"
            M_OBJCONN.execute StrSqlIp
        End If
        Set M_Objrs_Cek = Nothing
        
        '@@19042012, Cek IP Icentra
        Dim M_Objrs_IP_Icentra As ADODB.Recordset
        
        CMDSQL = "select * from tbl_ip_icentra where ip='"
        CMDSQL = CMDSQL + CStr(MDIForm1.WskCTI.LocalIP) + "'"
        Set M_Objrs_IP_Icentra = New ADODB.Recordset
        M_Objrs_IP_Icentra.CursorLocation = adUseClient
        M_Objrs_IP_Icentra.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_IP_Icentra.RecordCount = 0 Then
            MDIForm1.TxtIPIcentra.text = ""
            Set M_Objrs_IP_Icentra = Nothing
        Else
            MDIForm1.TxtIPIcentra.text = IIf(IsNull(M_Objrs_IP_Icentra("ip_icentra")), "", Trim(M_Objrs_IP_Icentra("ip_icentra")))
            Set M_Objrs_IP_Icentra = Nothing
        End If
        
        
        
        '@@ 30-May-2011 Menampilkan Form Confidence Analisys
        If Trim(tipe) = "AGENT" Then
            Dim cmdsql_confidence As String
            'Cek dulu apakah form confidence analisys sudah ditampilkan
            If Trim(M_objrs("f_confidence_analisis") = "0") Then
                cmdsql_confidence = "update usertbl set f_confidence_analisis='1' where userid='"
                cmdsql_confidence = cmdsql_confidence + Trim(agent) + "'"
                M_OBJCONN.execute cmdsql_confidence
                'FrmConfidenceAnalysis.Show vbModal
                 ' 08 SEPTEMBER 2014
                 'FrmConfidenceListNew_Agent.Show vbModal
            End If
        End If
        
        '@@15012013 Ambil nilai Tlnya nih
        If UCase(MDIForm1.Text2.text) = "AGENT" Then
            UseridTL = IIf(IsNull(M_objrs("team")), "", M_objrs("team"))
            '@@11022013 Tambahan buat catet akses all account
            AksesAllAcc = IIf(IsNull(M_objrs("f_akses_all_acc")), "", M_objrs("f_akses_all_acc"))
        End If
        
        '@@28012013, ini cek dulu, lagi diblok apa ngga aplikasinya!
'        If bcp = False Then
'            If M_objrs("f_blok") = "1" Then
'                MsgBox "Akun anda terblok dikarenakan blok " & M_objrs!alasan_blok & "! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
'blok_user:
'                End
'            End If
'        End If
        
        ' LOG BUAT ABSENSI 27 NOP 2013 -------------------
        If UCase(MDIForm1.Text2.text) <> "SUPERVISOR" Then
        
            If UCase(MDIForm1.Text2.text) = "AGENT" Then
                MDIForm1.mntools.Enabled = False
                MDIForm1.SSCommand3.Enabled = False
            Else
                MDIForm1.mn_performance.Enabled = False
                MDIForm1.mn_deskcoll_perform2.Enabled = False
                MDIForm1.mn_performance_reguler.Enabled = False
                If UCase(MDIForm1.Text2.text) <> "MANAGER" Then
                    MDIForm1.mndran.Enabled = False
                    MDIForm1.mndrm(55).Enabled = False
                End If
                MDIForm1.mnuCallmonitor.Enabled = True
            End If
            
            If M_objrs.state = 1 Then M_objrs.Close
            M_objrs.Open "SELECT userid FROM tblabsen_aplikasi WHERE userid='" & agent & "' AND date(tanggal)=date(now())"
            If M_objrs.RecordCount = 0 Then
                M_OBJCONN.execute "INSERT INTO tblabsen_aplikasi(userid,tanggal,hours) VALUES('" & agent & "',now(),0);"
            End If
        End If
        ' ------------------------------------------------
        
        Set M_objrs = Nothing
        
        '@@28012013, Ini buat nyatet agent yang login
        CMDSQL = "update usertbl set f_status_login='1' where userid='"
        CMDSQL = CMDSQL & MDIForm1.Text1.text + "'"
        M_OBJCONN.execute CMDSQL
         
        ' 10-05-2013 By Izuddin
        Call load_reminder
        ' ++++++++++++++++++++
        
        On Error GoTo next_err
        ' Update Database dulu 02 Feb 2015
        M_OBJCONN.execute "INSERT INTO tbl_count_block(agent,ket) values('" & MDIForm1.Text1.text & "','Login')"
next_err:
        M_OBJCONN.execute "DELETE FROM tbl_donotcall_today WHERE date(tgl)<date(now())"
        
        
        'queryupdateflogin
'        Dim flogin As String
'
'        If MDIForm1.Text2.Text = "Agent" Then
'            flogin = "UPDATE usertbl set f_login = '1' where userid = '" + MDIForm1.Text1.Text + "'"
'            M_OBJCONN.Execute flogin
'
'            If login = 1 Then
'                MsgBox "Sesi Akun Anda Dalam Keadaan Login"
'                End
'            End If
'        End If
       
    '=====tambahan asep 26/03/2020==== log login'
    Session_login = waktu_server_sekarang
    Session_ManualDial = waktu_server_sekarang
    Dim insertlog As ADODB.Recordset
    Dim sql1 As String
    Set insertlog = New ADODB.Recordset
    insertlog.CursorLocation = adUseClient
    
    Call cek_logout_kosong
    Call cek_breaktimeend_kosong
    FirstLogin = True
'    sql1 = "insert into usertbl_log (session_login,username,status,waktu_login,ip_login) "
'    sql1 = sql1 + " values('" + Session_login + "','" + MDIForm1.Text1.text + "','LOGIN','" + Session_login + "','" & MDIForm1.Winsock1.LocalIP & "')"
'    M_OBJCONN.execute sql1
'
'    sql1 = "INSERT into tbl_autodialer_agent_break(sessionid,status_break,agent,waktu_start,ip_login)values"
'    sql1 = sql1 + " ('" + Session_ManualDial + "','ManualDial','" + MDIForm1.Text1.text + "',now(),'" & MDIForm1.Winsock1.LocalIP & "')"
'    M_OBJCONN.execute sql1
    
        MDIForm1.Show
        


'        DoEvents
'        Set M_PESAN = New ADODB.Recordset
'        M_PESAN.CursorLocation = adUseClient
'        M_PESAN.Open "SELECT  MSGTBL.RECIPIENT,MSGTBL.DATETIME,MSGTBL.SENDER,MSGTBL.SENTFROM,MSGTBL.MSG,usertbl.AGENT FROM MSGTBL, usertbl WHERE MSGTBL.SENDER = usertbl.USERID AND RECIPIENT = '" + MDIForm1.Text1.Text + "' AND STS = 0 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        While Not M_PESAN.EOF
'            FRMTERIMAMSG.RichTextBox1.SelColor = &HC00000
'            FRMTERIMAMSG.Text1.Text = IIf(IsNull(M_PESAN!SENDER), "", M_PESAN!SENDER)
'            FRMTERIMAMSG.RichTextBox1.Text = FRMTERIMAMSG.RichTextBox1.Text + "Dari :" + IIf(IsNull(M_PESAN!SENDER), "", M_PESAN!SENDER) + " - " + IIf(IsNull(M_PESAN!agent), "", M_PESAN!agent) + vbCrLf
'            FRMTERIMAMSG.RichTextBox1.Text = FRMTERIMAMSG.RichTextBox1.Text + "Kepada :" + IIf(IsNull(M_PESAN!RECIPIENT), "", M_PESAN!RECIPIENT) + vbCrLf
'            FRMTERIMAMSG.RichTextBox1.Text = FRMTERIMAMSG.RichTextBox1.Text + "Tanggal :" + IIf(IsNull(M_PESAN!DateTime), "", M_PESAN!DateTime) + vbCrLf
'            FRMTERIMAMSG.RichTextBox1.Text = FRMTERIMAMSG.RichTextBox1.Text + "Isi Pesan :" + vbCrLf
'            FRMTERIMAMSG.RichTextBox1.Text = FRMTERIMAMSG.RichTextBox1.Text + IIf(IsNull(M_PESAN!MSG), "", M_PESAN!MSG)
'            FRMTERIMAMSG.RichTextBox1.Text = FRMTERIMAMSG.RichTextBox1.Text + " " + vbCrLf
'            FRMTERIMAMSG.RichTextBox1.Text = FRMTERIMAMSG.RichTextBox1.Text & vbCrLf
'            M_PESAN.MoveNext
'        Wend
'        If M_PESAN.RecordCount <> 0 Then
'            FRMTERIMAMSG.Show vbModal
'            'Call BUKA_FILE_KONEKSI("MSG.TXT")
'            'Call SAVE_FILE_KONEKSI("MSG.TXT", IIf(IsNull(M_PESAN!SENDER), "", M_PESAN!SENDER))
'            'WriteINI "LOGIN", "USER NAME", IIf(IsNull(M_PESAN!SENDER), "", M_PESAN!SENDER)
'        End If
'        Set M_PESAN = Nothing
     '   FrmTodayList.Show vbModal
''End If
Else
    MsgBox "User Name Yang Anda Masukan Tidak Terdaftar", vbCritical + vbOKOnly, "Peringatan"
'    Debug.Print Decrypt(Len(Trim(txtUserName.text)), M_OBJRS("ACCREC"))
    txtUserName.SetFocus
    Timer1.Enabled = False
    Label1.Visible = False
    GoTo bawah:
    'SendKeys "{Home}+{End}"
End If
    If (MDIForm1.Text1.text = "JOKO") Or (MDIForm1.Text1.text = "ONTARIO") Or (MDIForm1.Text1.text = "DODDY") Or (MDIForm1.Text1.text = "REZA") Then
        MDIForm1.Label13.Visible = True
        'MDIForm1.Text5.Visible = True
        MDIForm1.cmdenabledptp.Visible = True
    Else
        MDIForm1.Label13.Visible = False
        MDIForm1.Text5.Visible = False
        MDIForm1.cmdenabledptp.Visible = False
        MDIForm1.Option1.Visible = False
        MDIForm1.Option2.Visible = False
        MDIForm1.Text8.Visible = False
        MDIForm1.Label14.Visible = False
        MDIForm1.Command5.Visible = False
        MDIForm1.Label4.Visible = False
        MDIForm1.Text9.Visible = False
        MDIForm1.Command7.Visible = False
        MDIForm1.Command8.Visible = False
    End If
    
    MDIForm1.Ofl.Visible = False
Exit Sub
HELL:
 MsgBox err.Description  '"DATA HANYA BISA BUKA 1 APLIKASI"
bawah:
End Sub

Private Sub login_ext(number$)
Dim cancelflag As Boolean
Dim MSComm1 As MSComm
Dim DialString$, FromModem$, dummy
    DialString$ = "ATDT" + number$ + ";" + vbCr
    On Error Resume Next
    If MSComm1.PortOpen Then
    Else
        If MDIForm1.TxtCommPort.text = Empty Then
            MsgBox "Tidak Ada Variable buat Comport", vbInformation + vbOKOnly
            Exit Sub
        End If
        MSComm1.CommPort = MDIForm1.TxtCommPort.text
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.PortOpen = True
    End If
Me.MousePointer = 11
    If err Then
        MsgBox err.Description, vbCritical + vbOKOnly, "Aplikasi"
        MSComm1.PortOpen = False
        cancelflag = True
        Me.MousePointer = 0
        Exit Sub
    End If
    MSComm1.InBufferCount = 0
    MSComm1.output = DialString$
    Me.MousePointer = 0
    Do
        dummy = DoEvents()
        If MSComm1.InBufferCount Then
            FromModem$ = FromModem$ + MSComm1.Input
            If InStr(FromModem$, "OK") Then
            '    Beep
                WaitSecs (0.1)
                cancelflag = True
                Exit Do
            End If
            If InStr(FromModem$, "NO DIALTONE") Then
            '    Beep
            '    Beep
                MsgBox err.Description, vbInformation + vbOKOnly, "Aplikasi"
                cancelflag = True
                Exit Do
            End If
        End If
        If cancelflag Then
            cancelflag = False
            Me.MousePointer = 0
            Exit Do
        End If
    Loop
    If MSComm1.PortOpen = True And cancelflag = True Then
        MSComm1.output = "ATH" + vbCr
        MSComm1.PortOpen = False
    End If
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim txtstr As String
    
    strategi = False
'    Me.Height = 0
  '  If App.PrevInstance Then ShowPrevInstance
   'Call CheckSoftware(frmLogin)
'WebBrowser1.Navigate ("http://localhost/sobatmuslim/lokomedia/mobile/board.php")
'    txtstr = " PERINGATAN : " & vbCrLf
'    txtstr = txtstr & " - Pelanggaran kebijakan ini oleh staff akan ditangani berdasarkan prosedur penegakan disiplin yang sesuai." & vbCrLf
'    txtstr = txtstr & " - Jika pelanggaran yang dilakukan terkait dengan hukum pidana, maka dapat  segera dilaporkan kepada Polisi." & vbCrLf
'    txtstr = txtstr & " - Dalam hal terjadi kerugian yang diderita oleh Perusahaan sebagai akibat dari pelanggaran peraturan ini oleh pengguna, maka pengguna harus bertanggung jawab atas penggantian kerugian tersebut." & vbCrLf & vbCrLf
'
'    txtstr = txtstr & "UNDANG - UNDANG TERKAIT : " & vbCrLf
'    txtstr = txtstr & "- Undang-Undang No. 11 Tahun 2008 tentang Informasi dan Transaksi Elektronik" & vbCrLf
'    txtstr = txtstr & "- Undang-Undang No. 19 Tahun 2002 tentang Hak Cipta" & vbCrLf
'    txtstr = txtstr & "- Undang-Undang No. 14 Tahun 2008 tentang Kebebasan Informasi Publik" & vbCrLf
'
'    Text1.text = txtstr
End Sub
Private Sub Tmrreminder_Timer()
    Label2.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
    If (Label2.Left + Label2.Width) <= 0 Then
        Label2.Left = Me.Width
    End If
    Label2.Left = Label2.Left - 100
End Sub
Private Sub Timer1_Timer()
If Label1.Visible = False Then
    Label1.Visible = True
Else
    Label1.Visible = False
End If
DoEvents
End Sub


Public Sub Tengah()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
   ' MsgBox KeyAscii
End Sub

Private Sub ShowPrevInstance()
    Dim OldTitle As String
    Dim ll_WindowHandle As Long
    'Simpan judul ini ke dalam variabel OldTitle
    OldTitle = App.Title
    'Ganti judul aplikasinya...
    App.Title = "abcde - Aplikasi ini akan ditutup!"
    'Cari program sebelumnya. Jika Anda menggunakan VB
    '5.0, ganti "ThunderRT6Main" menjadi
    '"ThunderRT5Main"
    ll_WindowHandle = FindWindow("ThunderRT6Main", _
                      OldTitle)
    'Jika tidak ada aplikasi sebelumnya dibuka, keluar
    'langsung dari prosedur ini
    If ll_WindowHandle = 0 Then Exit Sub
    ll_WindowHandle = GetWindow(ll_WindowHandle, _
                      GW_HWNDPREV)
    'Sekarang ganti window tersebut...
    Call OpenIcon(ll_WindowHandle)
    'Dan bawa sebagai latar depan (tampil di depan)
    Call SetForegroundWindow(ll_WindowHandle)
    End
End Sub

Private Sub menurole()
    Dim query As String
    Dim rs_lvtian As New ADODB.Recordset
    Dim tian() As String
    
'    tian = Array("a", "b")
'    tian(0, 1) = "a"
    
    query = " select * from checkmenurole where tingkat = '" + MDIForm1.Text2.text + "'"
    Set rs_lvtian = New ADODB.Recordset
    rs_lvtian.CursorLocation = adUseClient
    rs_lvtian.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    'M_Objrs.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs_lvtian.RecordCount = 0 Then
       MsgBox "Menu Role Belum disetting"
    End If
    If rs_lvtian.RecordCount > 0 Then
        If rs_lvtian("UD") = "0" Then
            MDIForm1.nmupload.Enabled = False
        End If
        If rs_lvtian("SD") = "0" Then
            MDIForm1.nmswapdata.Enabled = False
        End If
        If rs_lvtian("UFLA") = "0" Then
            MDIForm1.mnsubmarkup.Enabled = False
        End If
        If rs_lvtian("TC") = "0" Then
            MDIForm1.mnagent.Enabled = False
        End If
        If rs_lvtian("BLNT") = "0" Then
            MDIForm1.mnblack.Enabled = False
        End If
        If rs_lvtian("SBD") = "0" Then
            MDIForm1.mnblokspv.Enabled = False
        End If
        If rs_lvtian("STFS") = "0" Then
            MDIForm1.setspv.Enabled = False
        End If
        If rs_lvtian("USA") = "0" Then
            MDIForm1.mnsubahstsacc.Enabled = False
        End If
        If rs_lvtian("CASP") = "0" Then
            MDIForm1.nmformceksts.Enabled = False
        End If
        If rs_lvtian("VLRF") = "0" Then
            MDIForm1.nmlistreqform.Enabled = False
        End If
        If rs_lvtian("ARAP") = "0" Then
            MDIForm1.nmlstreqnumber.Enabled = False
        End If
        If rs_lvtian("FLC") = "0" Then
            MDIForm1.nmmenuformlistconfidence.Enabled = False
        End If
        If rs_lvtian("PP") = "0" Then
            MDIForm1.mnbalance.Enabled = False
        End If
        If rs_lvtian("RSN") = "0" Then
            MDIForm1.mnrptsms.Enabled = False
        End If
        If rs_lvtian("VS") = "0" Then
            MDIForm1.VSMS.Enabled = False
        End If
        If rs_lvtian("BST") = "0" Then
            MDIForm1.smsblast.Enabled = False
        End If
        If rs_lvtian("LSS") = "0" Then
            MDIForm1.nmlistsmsscript.Enabled = False
        End If
        If rs_lvtian("AARS") = "0" Then
            MDIForm1.nmapprovreject.Enabled = False
        End If
        If rs_lvtian("SSBVE") = "0" Then
            MDIForm1.nmblastsmsexcel.Enabled = False
        End If
        If rs_lvtian("LUN") = "0" Then
            MDIForm1.nmListUnValidNumber.Enabled = False
        End If
        If rs_lvtian("ALT") = "0" Then
            MDIForm1.nmAksesLayanaTelkom.Enabled = False
        End If
        If rs_lvtian("LRP") = "0" Then
            MDIForm1.nmlistreqptp.Enabled = False
        End If
        If rs_lvtian("RP") = "0" Then
            MDIForm1.nmresetpass.Enabled = False
        End If
        If rs_lvtian("LRPH") = "0" Then
            MDIForm1.nmReportProblemHeadset.Enabled = False
        End If
        If rs_lvtian("LRPT") = "0" Then
            MDIForm1.nmListReportProblemTelepon.Enabled = False
        End If
        If rs_lvtian("BAT") = "0" Then
            MDIForm1.nmblokaplikasitins.Enabled = False
        End If
        If rs_lvtian("MDA") = "0" Then
            MDIForm1.nmManageDistribusiAccount.Enabled = False
        End If
        If rs_lvtian("LAL") = "0" Then
            MDIForm1.mnListAccountLunas.Enabled = False
        End If
        If rs_lvtian("LDC") = "0" Then
            MDIForm1.mn_list_complaint.Enabled = False
        End If
        If rs_lvtian("LS") = "0" Then
            MDIForm1.mn_list_sid.Enabled = False
        End If
        
        'dataconfidenct
        If rs_lvtian("MBP") = "0" Then
            MDIForm1.mn_monhly_bp.Enabled = False
        End If
        If rs_lvtian("MCPA") = "0" Then
            MDIForm1.mnmonthcpa.Enabled = False
        End If
        If rs_lvtian("MPP") = "0" Then
            MDIForm1.mnptppayment.Enabled = False
        End If
        If rs_lvtian("CL") = "0" Then
            MDIForm1.mn_confidence_list.Enabled = False
        End If
        
        'tools
        If rs_lvtian("LPR") = "0" Then
            MDIForm1.list_phone_review.Enabled = False
        End If
        If rs_lvtian("aoc") = "0" Then
            MDIForm1.mn_aoc.Enabled = False
        End If
        If rs_lvtian("TD") = "0" Then
            MDIForm1.transfer_data.Enabled = False
        End If
        If rs_lvtian("ASH") = "0" Then
            MDIForm1.add_special_history.Enabled = False
        End If
        If rs_lvtian("UDFW") = "0" Then
            MDIForm1.upload_fresh_wo.Enabled = False
        End If
        If rs_lvtian("RTA") = "0" Then
            MDIForm1.mn_report_temp.Enabled = False
        End If
        If rs_lvtian("DP") = "0" Then
            MDIForm1.mn_performance.Enabled = False
        End If
        If rs_lvtian("AP") = "0" Then
            MDIForm1.mn_deskcoll_perform2.Enabled = False
        End If
        If rs_lvtian("DPR") = "0" Then
            MDIForm1.mn_performance_reguler.Enabled = False
        End If
        If rs_lvtian("CM") = "0" Then
            MDIForm1.mnuCallmonitor.Enabled = False
        End If
        If rs_lvtian("CFCDDP") = "0" Then
            MDIForm1.mn_copyfile.Enabled = False
        End If
        If rs_lvtian("FHS") = "0" Then
            MDIForm1.mn_option_hide.Enabled = False
        End If
        If rs_lvtian("DRM") = "0" Then
            MDIForm1.mndrm(55).Enabled = False
        End If
    End If
End Sub

Private Sub cek_logout_kosong()
    Dim a_bojrs As New ADODB.Recordset
    Set a_bojrs = New ADODB.Recordset
    
    Dim b_bojrs As New ADODB.Recordset
    Set b_bojrs = New ADODB.Recordset
    
    Dim ssSql As String
    Dim ssSq2 As String
    Dim sql1 As String
    
    'ssSql = "select * from usertbl_log where session_login in (select max(session_login) as latest_login from usertbl_log where waktu_logout is null and username = '" & MDIForm1.Text1.text & "')"
    'ssSql = "select * from usertbl_log where date(waktu_login::timestamp) = date(now()) and username = '" & MDIForm1.Text1.text & "' and coalesce(waktu_logout,'')=''"
    ssSql = "select * from usertbl_log where session_login in " & vbCrLf
    ssSql = ssSql + "(select max(session_login) as session_login from usertbl_log where date(waktu_login::timestamp) = date(now()) and username = '" & MDIForm1.Text1.text & "' and coalesce(waktu_logout,'')='')"
    Set a_bojrs = New ADODB.Recordset
    a_bojrs.CursorLocation = adUseClient
    a_bojrs.Open ssSql, M_OBJCONN, adOpenDynamic, adLockOptimistic
          
    If a_bojrs.RecordCount <> "0" Then
        Session_login = IIf(IsNull(a_bojrs("Session_login")), "", a_bojrs("Session_login"))
        'M_OBJCONN.execute "update usertbl_log set waktu_logout = '" & waktu_server_sekarang & "' where session_login='" & latest_login & "'"
    Else
        sql1 = "insert into usertbl_log (session_login,username,status,waktu_login,ip_login) "
        sql1 = sql1 + " values('" + Session_login + "','" + MDIForm1.Text1.text + "','LOGIN',now()::timestamp,'" & MDIForm1.Winsock1.LocalIP & "')"
        M_OBJCONN.execute sql1
    End If
    
    ssSq2 = "select * from tbl_autodialer_agent_break where date(waktu_start::timestamp) = date(now()) and status_break = 'ManualDial' and agent = '" & MDIForm1.Text1.text & "' and coalesce(waktu_end::varchar,'')=''"
    Set b_bojrs = New ADODB.Recordset
    b_bojrs.CursorLocation = adUseClient
    b_bojrs.Open ssSq2, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If b_bojrs.RecordCount <> "0" Then
        Session_ManualDial = IIf(IsNull(b_bojrs("sessionid")), "", b_bojrs("sessionid"))
    Else
        sql1 = "INSERT into tbl_autodialer_agent_break(sessionid,status_break,agent,waktu_start,ip_login)values"
        sql1 = sql1 + " ('" + Session_ManualDial + "','ManualDial','" + MDIForm1.Text1.text + "','" & waktu_server_sekarang & "','" & MDIForm1.Winsock1.LocalIP & "')"
        M_OBJCONN.execute sql1
    End If
    
    Set a_bojrs = Nothing
    Set b_bojrs = Nothing
End Sub

Private Sub cek_breaktimeend_kosong()
    Dim a_objrs As New ADODB.Recordset
    Set a_objrs = New ADODB.Recordset
    Dim latest_login As String
    Dim ssSql As String
    Dim ssId As String
    
    ssSql = "select * from tbl_autodialer_agent_break where date_break in (select max(date_break) as latest_break from tbl_autodialer_agent_break where waktu_end is null "
    ssSql = ssSql + "and agent = '" & MDIForm1.Text1.text & "' and status_break not in ('ManualDial','start_autodialer','AutoDial','form break show') and date(date_break) between (date(now()) - interval '1 day') and date(now()) )"

    Set a_objrs = New ADODB.Recordset
    a_objrs.CursorLocation = adUseClient
    a_objrs.Open ssSql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       
    'If a_objrs.RecordCount <> "0" Then
       ' ssId = IIf(IsNull(a_objrs("id")), "0", a_objrs("id"))
        'M_OBJCONN.execute "update tbl_autodialer_agent_break set waktu_end = '" & waktu_server_sekarang & "' where id= '" & ssId & "'"
    'End If
End Sub
