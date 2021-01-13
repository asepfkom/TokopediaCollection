VERSION 5.00
Begin VB.Form FrmSendSmsNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send SMS"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbOption 
      Height          =   315
      ItemData        =   "FrmSendSmsNew.frx":0000
      Left            =   240
      List            =   "FrmSendSmsNew.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtnm_agent 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox TxtSmsFreeStyle 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   1150
      MaxLength       =   320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   1320
      Width           =   3315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   1150
      MaxLength       =   320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2520
      Width           =   3495
   End
   Begin VB.ComboBox CmbSubOption 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2100
      Width           =   3405
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   1260
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Perubahan Sistem SMS, berdasarkan status data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label lblid 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label LblLayer 
      Caption         =   "1"
      Height          =   315
      Left            =   960
      TabIndex        =   20
      Top             =   4620
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Layer:"
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   4620
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Nama :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Custid :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Agent :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   4890
   End
   Begin VB.Label Label2 
      Caption         =   "Text :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Mobile No :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah :"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Option:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1710
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Sub option:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2070
      Width           =   975
   End
End
Attribute VB_Name = "FrmSendSmsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public awal As String
Public btsakhir As Integer
Dim d As Integer
Dim awalk As Integer
Dim akhirk As Integer
Dim AvgMarks(50, 50) As Double
Dim rowaray As Integer

Private Sub CmbOption_Change()
     '@@Jika sms free style
    If CmbOption.text = "Free SMS Style" Then
        CmbSubOption.text = "Free SMS Style"
        Text1.text = "[]"
    End If
End Sub

Private Sub get_optionid()
    Dim M_Objrs     As ADODB.Recordset
    Dim cmdsql      As String
    
    cmdsql = "SELECT id FROM tblscriptsms WHERE option='" & Trim(CmbOption.text) & "' " & _
            "AND suboption='" & Trim(CmbSubOption.text) & "'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        lblID.Caption = M_Objrs!ID
    End If
  
    Set M_Objrs = Nothing
End Sub

Private Sub tblscript()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    cmdsql = "select * from tblscriptsms "
    If FrmCC_Colection.cboaccount.text Like "*VL*" Or FrmCC_Colection.cboaccount.text Like "*PR-*" Then
        cmdsql = cmdsql + " where id in (1,4) "
    ElseIf FrmCC_Colection.cboaccount.text Like "*ON*" Then
        cmdsql = cmdsql + " where id in (2,5) "
    ElseIf FrmCC_Colection.cboaccount.text Like "*POP*" Or FrmCC_Colection.cboaccount.text Like "*POP-PROGRESS OF PAYMENT*" Or FrmCC_Colection.cboaccount.text Like "*PTP-POP*" Or FrmCC_Colection.cboaccount.text Like "*BP-POP*" Then
        cmdsql = cmdsql + " where id in (3,8) "
    ElseIf FrmCC_Colection.cboaccount.text Like "*PTP-NEW*" Or FrmCC_Colection.cboaccount.text Like "*PTP-NE*" Or FrmCC_Colection.cboaccount.text Like "*PTP-PO*" Or FrmCC_Colection.cboaccount.text Like "*PO-*" Then
        cmdsql = cmdsql + " where id in (2) "
    ElseIf FrmCC_Colection.cboaccount.text Like "*BP-*" Or FrmCC_Colection.cboaccount.text Like "*BP-NEW*" Then
        cmdsql = cmdsql + " where id in (3) "
    ElseIf FrmCC_Colection.cboaccount.text Like "*OS*" Or FrmCC_Colection.cboaccount.text = "" Then
        cmdsql = cmdsql + " where id in (6,7) "
    End If
    
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql & " order by id", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    CmbSubOption.CLEAR
    While Not M_Objrs.EOF
        CmbSubOption.AddItem M_Objrs("suboption")
        M_Objrs.MoveNext
    Wend
    
  
    Set M_Objrs = Nothing
End Sub

Private Sub CmbOption_Click()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    Text1.text = ""
    CmbSubOption.CLEAR
    
    cmdsql = "select * from tblscriptsms order by id"
    'cmdsql = cmdsql + CmbOption.Text + "'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
'    If M_Objrs.RecordCount = 0 Then
'        '@@Jika sms free style
'        If CmbOption.Text = "Free SMS Style" Then
'            Text1.Visible = False
'            Text1.Text = ""
'            TxtSmsFreeStyle.Visible = True
'            TxtSmsFreeStyle.Text = ""
'            TxtSmsFreeStyle.SetFocus
'        End If
'        Set M_Objrs = Nothing
'        Exit Sub
'    Else
'        Text1.Visible = True
'        TxtSmsFreeStyle.Visible = False
'        TxtSmsFreeStyle.Text = ""
'    End If
    
    CmbSubOption.CLEAR
    While Not M_Objrs.EOF
        CmbSubOption.AddItem M_Objrs("suboption")
        M_Objrs.MoveNext
    Wend
    
  
    Set M_Objrs = Nothing
End Sub

Private Sub CmbOption_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbSubOption_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    Text1.text = ""
'    cmdsql = "select * from tblscriptsms where option='"
'    cmdsql = cmdsql + CmbOption.Text + "' and suboption='"
'    cmdsql = cmdsql + CmbSubOption.Text + "'"

    cmdsql = "select * from tblscriptsms where suboption= '" + CmbSubOption.text + "' "

    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    

    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    Text1.text = Trim(M_Objrs("scriptsms"))
    
    rowaray = 0
    For i = 1 To Len(Text1.text)
    If Mid(Text1.text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
            
    ElseIf Mid(Text1.text, i, 1) = "]" Then
        akhirk = i
         AvgMarks(1, rowaray) = i
         rowaray = rowaray + 1
    End If
    Next i
    If CmbOption.text = "Free SMS Style" Then
        Text1.text = "[]"
        lblID.Caption = 0
    Else
        Call get_optionid
    End If
    
    Text1.text = Replace(Text1.text, "*agent*", txtnm_agent.text)
    Text1.text = Replace(Text1.text, "*cust*", Text4.text)
    
    Set M_Objrs = Nothing
End Sub

Private Sub CmbSubOption_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo1_Click()
If Text5 = "" Then
If Left(Combo1, 1) <> "0" Then
Text5.text = Text5.text & "021" & Combo1.text
Else
Text5.text = Text5.text & Combo1.text
End If
Else
If Left(Combo1, 1) <> "0" Then
Text5.text = Text5.text & ",021" & Combo1.text
Else
Text5.text = Text5.text & "," & Combo1.text
End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Dim teks1 As String
    Dim teks2 As String
    Dim fields() As String
    Dim banyaksms As Integer
    Dim pesan As String
    Dim aa As Integer
    Dim m_objrscekkuota As ADODB.Recordset
    Dim SisaSms As Integer
    Dim SisaSmsSekrg As Integer
    
    '@@ 09022011 Ini jika  free sms style
    If CmbOption.text = "Free SMS Style" Then
        
        teks1 = Replace(TxtSmsFreeStyle.text, "[", "")
        teks1 = Replace(teks1, "'", "")
        teks1 = Replace(teks1, "\", "")
        
        teks2 = Replace(teks1, "]", "")
        teks2 = Replace(teks2, "'", "")
        
        'cek data udah di simpen ke tabel receive apa belum??
        fields() = Split(Text5.text, ",")
        For i = 0 To UBound(fields)
            '@@ 09022011 - Ambil tanggal system
            cmdsqltglsys = "SELECT now() AS tglsystem"
            Set R_tglsys = New ADODB.Recordset
            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not R_tglsys.EOF
                TGLw = R_tglsys("tglsystem")
                TGLSERVERc = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
                R_tglsys.MoveNext
            Wend
            Set R_tglsys = Nothing
                
'            Cmdsql = "select * from request_sms where agent='" & Trim$(Text2) & "' and custid='" & Trim$(Text3) & "' " & _
'                    " and notelp='" & Trim$(fields(i)) & "' and date(tgl_kirim)=date(now())"
            cmdsql = "SELECT * FROM request_sms WHERE custid='" & Trim$(Text3) & "' " & _
                    " AND notelp='" & Trim$(fields(i)) & "' AND date(tgl_kirim)=date(now()) AND id_option=" & Val(lblID.Caption) & ""
            Set M_Objrs = New ADODB.Recordset
           
            If M_Objrs.state = 1 Then M_Objrs.Close
     
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
            If M_Objrs.RecordCount = 0 Then
                GoTo lanjut
'                'Jika yang login admin/spv tidak usah hitung kuota sms
'                If UCase(Trim(MDIForm1.Text1.Text)) = "ADMIN" Or UCase(Trim(MDIForm1.Text1.Text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text1.Text)) = "SUPERVISOR" Then
'                    GoTo lanjut
'                End If
'                'Cek dulu deh udah berapa kali kirim
'                Set m_objrscekkuota = New ADODB.Recordset
'                m_objrscekkuota.CursorLocation = adUseClient
'                Cmdsql = "select kuota_sms from usertbl where userid='"
'                Cmdsql = Cmdsql + Trim$(Text2) + "'"
'                m_objrscekkuota.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                SisaSms = Val(MDIForm1.KuotaSms) - Val(m_objrscekkuota("kuota_sms"))
'                If Val(m_objrscekkuota("kuota_sms")) >= Val(MDIForm1.KuotaSms) Then
'                    MsgBox "Kuota sms anda hari ini sudah habis! Max.sms/hari:" & MDIForm1.KuotaSms & ". Silahkan coba esok hari!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                    Set m_objrscekkuota = Nothing
'                End If
'
'                'Jika sisa sms lebih kecil dari sisa kuota maka batalkan
'                If SisaSms < Val(LblLayer.Caption) Then
'                    MsgBox "SMS yang anda kirim melebihi kuota! Sisa sms yang diperbolehkan hari ini:" & SisaSms & " layer(160 karakter)", vbOKOnly + vbInformation, "Informasi"
'                    Set m_objrscekkuota = Nothing
'                    Exit Sub
'                End If
'
'                'Nah di sini jika kuoat masih tersedia, kurangi deh kuota dari agent tersebut
'                SisaSmsSekrg = Val(LblLayer.Caption) + Val(m_objrscekkuota("kuota_sms"))
'                Cmdsql = "update usertbl set kuota_sms='"
'                Cmdsql = Cmdsql + CStr(SisaSmsSekrg) + "' where userid='"
'                Cmdsql = Cmdsql + Trim$(Text2) + "'"
'                M_OBJCONN.Execute Cmdsql
'                Set m_objrscekkuota = Nothing
lanjut:
                cmdsql = "INSERT INTO request_sms "
                cmdsql = cmdsql + " ( agent, custid,name,notelp,pesan,status,tgl_kirim,id_option)"
                cmdsql = cmdsql + " VALUES"
                cmdsql = cmdsql + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Trim$(teks2) + "', '0', '" + TGLSERVERc + "'," & Val(lblID.Caption) & ")"
                M_OBJCONN.Execute cmdsql
                MsgBox "Sms berhasil disimpan! Akan dikirim setelah di approve oleh SPV!", vbOKOnly + vbInformation, "Informasi"
            Else
                MsgBox "Anda sudah mengirim sms ke no:" & Trim$(fields(i)) & ". Sebelumnya. SMS gagal dikirim!", vbOKOnly + vbInformation, "Informasi"
            End If
        Next
    Else
        '@@ Ini jika sms yang sesuai dengan format yang sudah ditentukan
        teks1 = Replace(Text1.text, "[", "")
        teks2 = Replace(teks1, "]", "")
        teks2 = Replace(teks2, "'", "")
        
        'cek data udah di simpen ke tabel receive apa belum??
        fields() = Split(Text5.text, ",")
        For i = 0 To UBound(fields)
            '@@ 09022011 - Ambil tanggal system
            cmdsqltglsys = "SELECT now() AS tglsystem"
            Set R_tglsys = New ADODB.Recordset
            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not R_tglsys.EOF
                TGLw = R_tglsys("tglsystem")
                TGLSERVERc = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
                R_tglsys.MoveNext
            Wend
            Set R_tglsys = Nothing
                
'            Cmdsql = "select * from request_sms where agent='" & Trim$(Text2) & "' and custid='" & Trim$(Text3) & "' " & _
'                    " and notelp='" & Trim$(fields(i)) & "' and date(tgl_kirim)=date(now())"
            cmdsql = "SELECT * FROM request_sms WHERE custid='" & Trim$(Text3) & "' " & _
                    " AND notelp='" & Trim$(fields(i)) & "' AND date(tgl_kirim)=date(now()) AND id_option=" & Val(lblID.Caption) & ""
            Set M_Objrs = New ADODB.Recordset
           
            If M_Objrs.state = 1 Then M_Objrs.Close
     
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
            If M_Objrs.RecordCount = 0 Then
                cmdsql = "INSERT INTO request_sms_log "
                cmdsql = cmdsql + " ( agent, custid,name,notelp,pesan,status,tgl_kirim)"
                cmdsql = cmdsql + " VALUES"
                cmdsql = cmdsql + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Trim$(teks2) + "', '0', '" + TGLSERVERc + "')"
                M_OBJCONN.Execute cmdsql
                
                '@@ 09022011 Tambahan jika pesan lebih dari 160 karakter
                '@@ 11 May 2011 Permintaan SPV PIL, agar sms yang dikirim harus melalui approve
'                banyaksms = Ceiling(Val(Len(Trim$(fields(i)))) / 160)
'                For aa = 1 To banyaksms
'                    awalpesan = (160 * aa) - 160
'                    Pesan = Mid(Trim$(fields(i)), awalpesan + 1, 160)
'                    cmdsql = "insert into outbox (destinationnumber,textdecoded,senderid,creatorid) values ('"
'                    cmdsql = cmdsql + Trim(Pesan) + "','"
'                    cmdsql = cmdsql + Trim$(teks2) + "','phone1','"
'                    cmdsql = cmdsql + "PIL" + Trim$(Text3) + "-" + Trim$(Text2) + "')"
'                    M_OBJCONN1.Execute cmdsql
'                Next aa
'                MsgBox "SMS anda sudah terkirim!Sebanyak :" & banyaksms, vbOKOnly + vbInformation, "Informasi"
           
                cmdsql = "INSERT INTO request_sms "
                cmdsql = cmdsql + " ( agent, custid,name,notelp,pesan,status,tgl_kirim,id_option)"
                cmdsql = cmdsql + " VALUES"
                cmdsql = cmdsql + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Trim$(teks2) + "', '0', '" + TGLSERVERc + "'," & Val(lblID.Caption) & ")"
                M_OBJCONN.Execute cmdsql
           
                MsgBox "SMS anda telah disimpan dan akan dikirim setelah di approve oleh SPV!", vbOKOnly + vbInformation, "Informasi"
            Else
                MsgBox "Anda sudah mengirim sms ke no:" & Trim$(fields(i)) & ". Sebelumnya. SMS gagal dikirim!", vbOKOnly + vbInformation, "Informasi"
            End If
        Next
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Dim teks1 As String
    Dim teks2 As String
    
    teks1 = Replace(Text1.text, "[", "")
    teks2 = Replace(teks1, "]", "")
    
    MsgBox teks2
End Sub

Private Sub Form_Load()
Dim RSsms_send As ADODB.Recordset
Set RSsms_send = New ADODB.Recordset
Dim lst As listItem

On Error Resume Next

RSsms_send.CursorLocation = adUseClient
cmdsql = "SELECT btrim as no_tlp FROM ("
cmdsql = cmdsql + "    SELECT trim(mobileno) FROM mgm WHERE trim(mobileno) not in (select no_telp from tblblacklist) and custid        = '" + FrmCC_Colection.lblCustId + "' "
cmdsql = cmdsql + "    Union All"
cmdsql = cmdsql + "    SELECT trim(mobileno2) FROM mgm WHERE trim(mobileno2) not in (select no_telp from tblblacklist) and            custid = '" + FrmCC_Colection.lblCustId + "' "
'cmdsql = cmdsql + "    Union All"
'cmdsql = cmdsql + "    SELECT trim(mobilenoadd1) FROM mgm WHERE trim(mobilenoadd1) not in (select no_telp from tblblacklist) and       custid = '" + FrmCC_Colection.lblCustId + "' "
'cmdsql = cmdsql + "    Union All"
'cmdsql = cmdsql + "    SELECT trim(mobilenoadd2) FROM mgm WHERE trim(mobilenoadd2) not in (select no_telp from tblblacklist) and       custid = '" + FrmCC_Colection.lblCustId + "'
cmdsql = cmdsql + " ) a "
RSsms_send.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 
'While Not RSsms_send.EOF
'    Combo1.AddItem Replace(Trim(RSsms_send("no_tlp")), " ", "")
'    RSsms_send.MoveNext
'Wend
If FrmCC_Colection.Label22(2).Caption = 1 Then
    Combo1.text = FrmCC_Colection.tdbvalid.Value
    Combo1_Click
    Combo1.Enabled = False
Else
    While Not RSsms_send.EOF
        Combo1.AddItem Replace(Trim(RSsms_send("no_tlp")), " ", "")
        RSsms_send.MoveNext
    Wend
End If

Call tblscript


'RSsms_send.CursorLocation = adUseClient
'cmdsql = "Select * from mgm where custid='" + FrmCC_Colection.lblCustId + "'"
'RSsms_send.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'While Not RSsms_send.EOF
'
'    If (IsNull(RSsms_send("mobileno"))) Or RSsms_send("mobileno") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobileno")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobileno")), " ", "")
'        End If
'    End If
'
'    If (IsNull(RSsms_send("mobileno2"))) Or RSsms_send("mobileno2") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobileno2")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobileno2")), " ", "")
'        End If
'    End If
'
'    If (IsNull(RSsms_send("mobilenoadd1"))) Or RSsms_send("mobilenoadd1") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobilenoadd1")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobilenoadd1")), " ", "")
'        End If
'    End If
'
'    If (IsNull(RSsms_send("mobilenoadd2"))) Or RSsms_send("mobilenoadd2") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobilenoadd2")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobilenoadd2")), " ", "")
'        End If
'    End If


Set RSsms_send = Nothing

Text3 = FrmCC_Colection.lblCustId
Text4 = FrmCC_Colection.lblnama
Text2 = MDIForm1.Text1
txtnm_agent.text = MDIForm1.Text7.text

Load_Data_Option_SMSScript
End Sub


Private Sub Text1_Change()
Label6 = "Jumlah : " & Len(Text1)

'LblLayer.Caption = Ceiling(Val(Len(Trim(Text1.text))) / 160)

If Len(Text1) > 320 Then
    MsgBox "Hanya dapat mengirim sms sebanyak 320 Karakter"
End If
End Sub

Private Sub Load_Data_Option_SMSScript()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmdsql = "select distinct option from tblscriptsms"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    CmbOption.CLEAR
    'CmbOption.AddItem "Free SMS Style"
    While Not M_Objrs.EOF
        CmbOption.AddItem M_Objrs("option")
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim cek As Boolean
    cek = False
    For K = 0 To rowaray - 1
        Debug.Print Text1.SelStart
        update
        If Text1.SelStart >= AvgMarks(0, K) And Text1.SelStart < AvgMarks(1, K) Then
  
            If KeyAscii = vbKeyBack Then
                a = Mid(Text1.text, Text1.SelStart, 1)
                If a = "[" Or a = "]" Then
                    KeyAscii = 0
                End If
            End If
            cek = True
            Exit For
        End If
    Next K

    If cek = False Then
        KeyAscii = 0
    End If
End Sub

Public Sub update()
    Dim i As Integer
    rowaray = 0
    For i = 1 To Len(Text1.text)
        If Mid(Text1.text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
        ElseIf Mid(Text1.text, i, 1) = "]" Then
            akhirk = i
            AvgMarks(1, rowaray) = i
            rowaray = rowaray + 1
        End If
    Next i
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        MsgBox "Anda tidak dapat menggunakan klik kanan!", vbCritical + vbOKOnly, "Peringatan"
        Text1.text = ""
    End If
End Sub

Private Sub TxtSmsFreeStyle_Change()
    Label6 = "Jumlah : " & Len(TxtSmsFreeStyle.text)
    
    LblLayer.Caption = Ceiling(Val(Len(Trim(TxtSmsFreeStyle.text))) / 160)

    If Len(TxtSmsFreeStyle.text) > 320 Then
        MsgBox "Hanya dapat mengirim sms sebanyak 320 Karakter"
    End If
End Sub
'@@09022011 Fungsi buat membulatkan desimal
Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function

